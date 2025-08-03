import ExcelJS from 'exceljs';



// Column mappings
const COLUMN_MAPPINGS = {
  'Order Number': 'Order #',
  'Flat Number': 'Flat #',
  'Customer Mobile Number': 'Mobile No',
  'Confirmed Order': 'Cnf',
  'Product Name': 'Product Name',
  'Item Count': 'Qty',
  'Rate': 'Price',
  'Item total': 'I Tot',
  'Total Items': 'Total Items',
  'Payment Mode': 'Payment Mode',
  'Payment Status': 'Payment Status',
  'Total Amount': 'T Amt'
};

// Adjusted column widths to prevent ## displaying
const COLUMN_WIDTHS = {
  'Order #': 7,
  'Flat #': 7,
  'Mobile No': 16,
  'Cnf': 2,
  'Product Name': 35,
  'Qty': 2.5,
  'Price': 5.5,
  'I Tot': 5,        // Increased for numeric values
  'Total Items': 6,
  'Payment Mode': 5.75,
  'Payment Status': 5.75,
  'T Amt': 6,        // Increased for totals
};

// Column widths for Sheet2 (includes additional columns)
const COLUMN_WIDTHS_SHEET2 = {
  'Order #': 7,
  'Flat #': 7,
  'Mobile No': 16,
  'Cnf': 2,
  'Product Name': 35,
  'Qty': 2.5,
  'Price': 5.5,
  'I Tot': 5,
  'Total Items': 6,
  'Payment Mode': 5.75,
  'Payment Status': 5.75,
  'T Amt': 6,
  'Catalogue Group': 20,    // New column
  'Tax %': 8,               // New column
  'Tax Amount': 10          // New column
};

function extractNumberFromAddress(address) {
  const addressStr = String(address || '').trim();
  const match = addressStr.match(/\d+/); // Find first sequence of digits
  return match ? match[0] : '';
}

export async function filterExcel(filePath, custDataFilePath) {
  // Read the main workbook
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(filePath);
  
  // Read the customer data workbook
  const custDataWorkbook = new ExcelJS.Workbook();
  await custDataWorkbook.xlsx.readFile(custDataFilePath);
  
  const sheetName = 'Inquiries with order meta';
  const sheet = workbook.getWorksheet(sheetName);
  
  if (!sheet) {
    throw new Error(`Sheet '${sheetName}' not found in the uploaded file.`);
  }
  
  // Extract data
  const data = [];
  sheet.eachRow((row, rowNumber) => {
    if (rowNumber === 1) return; // Skip header
    
    const rowData = {};
    row.eachCell((cell, colNumber) => {
      const header = sheet.getRow(1).getCell(colNumber).value;
      rowData[header] = cell.value || '';
    });
    data.push(rowData);
  });

  // STEP 1: Filter orders which are neither "COMPLETED" nor "REJECTED"
// STEP 1: Filter orders which are neither "COMPLETED" nor "REJECTED"
// STEP 1: Filter orders which are neither "COMPLETED" nor "REJECTED"
const filteredData = data.filter(row => {
  const orderStatus = String(row['Order Status'] || '').toUpperCase().trim();
  return orderStatus !== 'COMPLETED' && orderStatus !== 'REJECTED';
});

// Group filtered rows by Order Number
const orderGroups = {};
filteredData.forEach(row => {
  const orderNum = row['Order Number'];
  if (!orderGroups[orderNum]) orderGroups[orderNum] = [];
  orderGroups[orderNum].push(row);
});

// Calculate Total Amount for each filtered order
const orderTotals = {};
for (const orderNum in orderGroups) {
  let sum = 0;
  orderGroups[orderNum].forEach(row => {
    const count = parseFloat(row['Item Count']) || 0;
    const discountedPrice = parseFloat(row['Discounted Price']) || 0;
    const regularPrice = parseFloat(row['Price']) || 0;
    const price = discountedPrice || regularPrice;
    sum += count * price;
  });
  orderTotals[orderNum] = Math.round(sum * 100) / 100;
}

// Parse Cust_Data sheet and build lookup
const custDataSheet = custDataWorkbook.getWorksheet('Cust_Data');
if (!custDataSheet) {
  throw new Error(`Sheet 'Cust_Data' not found in the customer data file.`);
}

const custLookup = {};
custDataSheet.eachRow((row, rowNumber) => {
  if (rowNumber === 1) return; // Skip header
  let mbNo = '';
  let flatNo = '';
  row.eachCell((cell, colNumber) => {
    const header = custDataSheet.getRow(1).getCell(colNumber).value;
    if (header === 'Mb No') mbNo = String(cell.value || '').trim();
    if (header === 'Flat No') flatNo = String(cell.value || '').trim();
  });
  if (mbNo) custLookup[mbNo] = flatNo;
});

// NEW: Address validation step
// NEW: Address validation step
const flaggedOrders = [];
const validOrders = [];

filteredData.forEach(row => {
  const mobileNo = String(row['Customer Mobile Number'] || '').trim();
  const shippingAddress = String(row['Shipping Address'] || '').trim();
  const lookupFlatNo = custLookup[mobileNo];
  
  if (lookupFlatNo) {
    // If no shipping address but mobile number exists in template, treat as valid
    if (!shippingAddress) {
      validOrders.push(row);
      return;
    }
    
    // Extract number from shipping address and flat number
    const addressNumber = extractNumberFromAddress(shippingAddress);
    const flatNumber = extractNumberFromAddress(lookupFlatNo);
    
    // Compare the extracted numbers
    if (addressNumber && flatNumber && addressNumber === flatNumber) {
      validOrders.push(row);
    } else {
      flaggedOrders.push(row);
    }
  } else {
    // If no lookup found, treat as valid (existing logic)
    validOrders.push(row);
  }
});

// ✅ NOW calculate totals for flagged orders (AFTER flaggedOrders is defined)
const flaggedOrderGroups = {};
flaggedOrders.forEach(row => {
  const orderNum = row['Order Number'];
  if (!flaggedOrderGroups[orderNum]) flaggedOrderGroups[orderNum] = [];
  flaggedOrderGroups[orderNum].push(row);
});

const flaggedOrderTotals = {};
const flaggedOrderItemTotals = {};
for (const orderNum in flaggedOrderGroups) {
  let sum = 0;
  let totalItems = 0;
  flaggedOrderGroups[orderNum].forEach(row => {
    const count = parseFloat(row['Item Count']) || 0;
    const discountedPrice = parseFloat(row['Discounted Price']) || 0;
    const regularPrice = parseFloat(row['Price']) || 0;
    const price = discountedPrice || regularPrice;
    sum += count * price;
    totalItems += count;
  });
  flaggedOrderTotals[orderNum] = Math.round(sum * 100) / 100;
  flaggedOrderItemTotals[orderNum] = totalItems;
}

// Continue with validOrders instead of filteredData
// Continue with validOrders instead of filteredData
const mainOrders = [];
const newNumOrders = [];

for (const row of validOrders) {
  const mobileNo = String(row['Customer Mobile Number'] || '').trim();
  
  // If mobile number exists in custLookup, put in mainOrders
  // If mobile number does NOT exist in custLookup, put in newNumOrders
  if (custLookup[mobileNo]) {
    mainOrders.push(row);
  } else {
    newNumOrders.push(row);
  }
}

  // 4. For mainOrders, update Flat # using lookup
  mainOrders.forEach(row => {
    const mobileNo = String(row['Customer Mobile Number'] || '').trim();
    if (custLookup[mobileNo]) {
      row['Flat Number'] = custLookup[mobileNo];
    }
  });

  // 5. Group all orders by Order # and sort by Flat #
  // 5. Group all orders by Order # and sort by Flat #'s first letter
// 5. Group all orders by Order # and sort by Flat # (floor-wise)
function groupAndSort(orders, moveEmptyToBottom = false) {
  // Helper function to parse flat number into components
  function parseFlatNumber(flatNo) {
    const str = String(flatNo || '').trim();
    if (!str) return { tower: '', floor: 0, apt: 0, isEmpty: true };
    
    // Extract tower letter(s) and number
    const match = str.match(/^([A-Z]+)(\d+)$/i);
    if (!match) return { tower: str, floor: 0, apt: 0, isEmpty: !str };
    
    const tower = match[1].toUpperCase();
    const numberPart = match[2];
    
    // Parse floor and apartment based on length
    let floor, apt;
    if (numberPart.length === 3) {
      // Format: X01 (floor 0, apt 1)
      floor = parseInt(numberPart.substring(0, 1));
      apt = parseInt(numberPart.substring(1));
    } else if (numberPart.length === 4) {
      // Format: 1007 (floor 10, apt 7)
      floor = parseInt(numberPart.substring(0, 2));
      apt = parseInt(numberPart.substring(2));
    } else {
      // Fallback for other formats
      floor = parseInt(numberPart);
      apt = 0;
    }
    
    return { tower, floor, apt, isEmpty: false };
  }
  
  // Group by Order Number
  const groups = {};
  orders.forEach(row => {
    const orderNum = row['Order Number'];
    if (!groups[orderNum]) groups[orderNum] = [];
    groups[orderNum].push(row);
  });
  
  // Sort groups by the parsed flat number of the first item
  const sortedGroups = Object.entries(groups).sort((a, b) => {
    const flatA = parseFlatNumber(a[1][0]['Flat Number']);
    const flatB = parseFlatNumber(b[1][0]['Flat Number']);
    
    // If moveEmptyToBottom is true, put empty flats at the bottom
    if (moveEmptyToBottom) {
      if (flatA.isEmpty && !flatB.isEmpty) return 1;
      if (!flatA.isEmpty && flatB.isEmpty) return -1;
      if (flatA.isEmpty && flatB.isEmpty) return 0;
    }
    
    // First sort by tower
    if (flatA.tower !== flatB.tower) {
      return flatA.tower.localeCompare(flatB.tower);
    }
    // Then by floor
    if (flatA.floor !== flatB.floor) {
      return flatA.floor - flatB.floor;
    }
    // Finally by apartment
    return flatA.apt - flatB.apt;
  });
  
  // Flatten the sorted groups
  return sortedGroups.flatMap(([orderNum, group]) => group);
}
const sortedMainOrders = groupAndSort(mainOrders, true);
  const sortedNewNumOrders = groupAndSort(newNumOrders);
  // Calculate total items per order
const orderItemTotals = {};
for (const orderNum in orderGroups) {
  let totalItems = 0;
  orderGroups[orderNum].forEach(row => {
    const itemCount = parseFloat(row['Item Count']) || 0;
    totalItems += itemCount;
  });
  orderItemTotals[orderNum] = totalItems;
}


  // 6. Transform for output (reuse your transformation logic)
  function transformRows(rows, customOrderTotals = null, customOrderItemTotals = null) {
    return rows.map(row => {
      const itemCount = parseFloat(row['Item Count']) || 0;
      const discountedPrice = parseFloat(row['Discounted Price']) || 0;
      const regularPrice = parseFloat(row['Price']) || 0;
      const rate = discountedPrice || regularPrice;
      const itemTotal = Math.round((itemCount * rate) * 100) / 100;
      const orderNum = row['Order Number'];
      let confirmedOrder = String(row['Confirmed Order']).toUpperCase().trim();
      confirmedOrder = confirmedOrder === 'TRUE' ? 'T' : 'F';
      let paymentMode = '';
      let paymentStatus = 'Due';
      const originalPaymentMode = String(row['Payment Mode'] || '').trim();
      const originalPaymentStatus = String(row['Payment Status'] || '').trim();
      if (originalPaymentMode.toLowerCase() === 'phonepe' && originalPaymentStatus.toUpperCase() === 'SUCCESSFUL') {
        paymentMode = 'ONL';
        paymentStatus = 'Paid';
      }
      
      // Use custom totals if provided (for flagged orders), otherwise use global totals
      const totalAmount = customOrderTotals ? customOrderTotals[orderNum] : orderTotals[orderNum];
      const totalItems = customOrderItemTotals ? customOrderItemTotals[orderNum] : orderItemTotals[orderNum];
      
      return {
        'Order #': row['Order Number'],
        'Flat #': row['Flat Number'],
        'Mobile No': row['Customer Mobile Number'],
        'Cnf': confirmedOrder,
        'Product Name': row['Product Name'],
        'Qty': itemCount,
        'Price': rate,
        'I Tot': itemTotal,
        'Total Items': totalItems || 0,
        'Payment Mode': paymentMode,
        'Payment Status': paymentStatus,
        'T Amt': totalAmount,
      };
    });
  }

  // Transform function for Sheet2 with additional columns
function transformRowsSheet2(rows, customOrderTotals = null, customOrderItemTotals = null) {
  return rows.map(row => {
    const itemCount = parseFloat(row['Item Count']) || 0;
    const discountedPrice = parseFloat(row['Discounted Price']) || 0;
    const regularPrice = parseFloat(row['Price']) || 0;
    const rate = discountedPrice || regularPrice;
    const itemTotal = Math.round((itemCount * rate) * 100) / 100;
    const orderNum = row['Order Number'];
    let confirmedOrder = String(row['Confirmed Order']).toUpperCase().trim();
    confirmedOrder = confirmedOrder === 'TRUE' ? 'T' : 'F';
    let paymentMode = '';
    let paymentStatus = 'Due';
    const originalPaymentMode = String(row['Payment Mode'] || '').trim();
    const originalPaymentStatus = String(row['Payment Status'] || '').trim();
    if (originalPaymentMode.toLowerCase() === 'phonepe' && originalPaymentStatus.toUpperCase() === 'SUCCESSFUL') {
      paymentMode = 'ONL';
      paymentStatus = 'Paid';
    }
    
    // Use custom totals if provided (for flagged orders), otherwise use global totals
    const totalAmount = customOrderTotals ? customOrderTotals[orderNum] : orderTotals[orderNum];
    const totalItems = customOrderItemTotals ? customOrderItemTotals[orderNum] : orderItemTotals[orderNum];
    
    return {
      'Order #': row['Order Number'],
      'Flat #': row['Flat Number'],
      'Mobile No': row['Customer Mobile Number'],
      'Cnf': confirmedOrder,
      'Product Name': row['Product Name'],
      'Qty': itemCount,
      'Price': rate,
      'I Tot': itemTotal,
      'Total Items': totalItems || 0,
      'Payment Mode': paymentMode,
      'Payment Status': paymentStatus,
      'T Amt': totalAmount,
      'Catalogue Group': row['Catalogue Group'] || '',     // New column
      'Tax %': row['Tax %'] || '',                         // New column
      'Tax Amount': row['Tax Amount'] || ''                // New column
    };
  });
}
  const transformedMainOrders = transformRows(sortedMainOrders);
  const transformedNewNumOrders = transformRows(sortedNewNumOrders);
  
  // Combine all transformed orders for further processing
  const allTransformedOrders = [...transformedMainOrders, ...transformedNewNumOrders];
  
  // 7. Write to sheets
// 7. Write to sheets
// 7. Write to sheets
const newWorkbook = new ExcelJS.Workbook();

// Flagged_Add sheet (create first if there are flagged orders)
if (flaggedOrders.length > 0) {
  const flaggedSheet = newWorkbook.addWorksheet('Flagged_Add');
  const transformedFlaggedOrders = transformRows(flaggedOrders, flaggedOrderTotals, flaggedOrderItemTotals);
  await addDataToSheet(flaggedSheet, transformedFlaggedOrders);
}

// Main sheet (Sheet1)
const mainSheet = newWorkbook.addWorksheet('Sheet1');
await addDataToSheet(mainSheet, transformedMainOrders);

// Sheet2 with additional columns
const allValidOrders = [...sortedMainOrders, ...sortedNewNumOrders];
const transformedSheet2Orders = transformRowsSheet2(allValidOrders);
const sheet2 = newWorkbook.addWorksheet('Sheet2');
await addDataToSheet(sheet2, transformedSheet2Orders, false, true); // true for useSheet2Columns

// New_Num sheet
if (transformedNewNumOrders.length > 0) {
  const newNumSheet = newWorkbook.addWorksheet('New_Num');
  await addDataToSheet(newNumSheet, transformedNewNumOrders);
}
  
  // Create tower sheets
  const towers = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'J', 'K', 'L', 'M', 'N', 'P'];
  
  for (const tower of towers) {
    const towerData = allTransformedOrders.filter(row => { // ✅ FIXED: Using allTransformedOrders
      const flatNo = String(row['Flat #'] || '');
      return flatNo.toUpperCase().startsWith(tower);
    });
    
    if (towerData.length > 0) {
      const towerSheet = newWorkbook.addWorksheet(`Tower ${tower}`);
      await addDataToSheet(towerSheet, towerData, true);
    }
  }

  // Handle customer details
  await handleCustomerDetails(workbook, newWorkbook, allTransformedOrders);

  // Write to buffer
  const buffer = await newWorkbook.xlsx.writeBuffer();
  return buffer;
}

async function addDataToSheet(worksheet, data, addBlankRows = false, useSheet2Columns = false) {
  // Choose which column set to use
  const columnWidths = useSheet2Columns ? COLUMN_WIDTHS_SHEET2 : COLUMN_WIDTHS;
  const columns = Object.keys(columnWidths);
  
  // Set up columns with proper widths
  worksheet.columns = columns.map(col => ({
    header: col,
    key: col,
    width: columnWidths[col]
  }));

  // Style the header row
  const headerRow = worksheet.getRow(1);
  headerRow.height = 78;
  
  // Apply header formatting
  columns.forEach((col, index) => {
    const cell = headerRow.getCell(index + 1);
    cell.font = { bold: true, size: 12 };
    cell.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
    cell.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'FFFF00' } // Yellow background
    };
    cell.border = {
      top: { style: 'thin' },
      bottom: { style: 'thin' },
      left: { style: 'thin' },
      right: { style: 'thin' }
    };
  });

  // Add data rows
  let lastFlatNo = null;
  
  data.forEach(row => {
    // Add blank row when flat number changes (for tower sheets)
    if (addBlankRows && lastFlatNo && lastFlatNo !== row['Flat #']) {
      const blankRow = worksheet.addRow({});
      blankRow.height = 15; // Standard row height
      // Apply left alignment to blank row
      blankRow.eachCell(cell => {
        cell.alignment = { horizontal: 'left' };
      });
    }
    
    const dataRow = worksheet.addRow(row);
    dataRow.height = 15; // Standard row height
    
    // Apply formatting to data cells
    columns.forEach((col, index) => {
      const cell = dataRow.getCell(index + 1);
      
      // Base font setting
      let font = { size: 12 };
      
      // Add left alignment
      cell.alignment = { horizontal: 'left' };
      
      // Only add borders
      cell.border = {
        top: { style: 'thin' },
        bottom: { style: 'thin' },
        left: { style: 'thin' },
        right: { style: 'thin' }
      };
      
      // Format numeric columns
      if (col === 'Qty' || col === 'Price' || col === 'I Tot' || col === 'T Amt' || col === 'Total Items' || col === 'Tax %' || col === 'Tax Amount') {
        cell.numFmt = '#,##0'; // Number format
        
        if (col === 'Qty' && parseFloat(cell.value) > 1) {
          font = { bold: true, size: 12, color: { argb: 'FF008000' } }; // Green
        }
      }
      
      if (col === 'Payment Status' && cell.value === 'Due') {
        font = { bold: true, size: 12, color: { argb: 'FFFF0000' } }; // Red
      }
      
      cell.font = font;
    });
    
    lastFlatNo = row['Flat #'];
  });
}

async function handleCustomerDetails(originalWorkbook, newWorkbook, filteredRows) {
  const customerSheetName = 'Cust_Data';
  const customerSheet = originalWorkbook.getWorksheet(customerSheetName);
  
  let customerData = [];
  
  if (customerSheet) {
    // Extract existing customer data
    customerSheet.eachRow((row, rowNumber) => {
      if (rowNumber === 1) return; // Skip header
      
      const customer = {};
      row.eachCell((cell, colNumber) => {
        const header = customerSheet.getRow(1).getCell(colNumber).value;
        customer[header] = cell.value || '';
      });
      customerData.push(customer);
    });
  }
  
  // Extract existing mobile numbers
  const existingMobileNumbers = new Set(
    customerData.map(row => row['Customer Mobile Number'] || row['Mobile Number'])
  );
  
  // Find new customers
  const newCustomers = [];
  const processedMobileNumbers = new Set();
  
  filteredRows.forEach(row => {
    const mobileNo = row['Mobile No'];
    const flatNo = row['Flat #'];
    
    if (mobileNo && !existingMobileNumbers.has(mobileNo) && 
        !processedMobileNumbers.has(mobileNo) && flatNo) {
      newCustomers.push({
        'Customer Mobile Number': mobileNo,
        'Flat Number': flatNo
      });
      processedMobileNumbers.add(mobileNo);
    }
  });
  
  if (customerData.length > 0 || newCustomers.length > 0) {
    const allCustomers = [...customerData, ...newCustomers];
    const newCustomerSheet = newWorkbook.addWorksheet(customerSheetName);
    
    // Set up columns
    newCustomerSheet.columns = [
      { header: 'Customer Mobile Number', key: 'Customer Mobile Number', width: 20 },
      { header: 'Flat Number', key: 'Flat Number', width: 15 }
    ];
    
    // Style ONLY the header row
    const headerRow = newCustomerSheet.getRow(1);
    headerRow.height = 30;
    headerRow.eachCell(cell => {
      cell.font = { bold: true, size: 12 };
      cell.alignment = { vertical: 'middle', horizontal: 'center' };
      cell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FFFF00' } 
      };
      cell.border = {
        top: { style: 'thin' },
        bottom: { style: 'thin' },
        left: { style: 'thin' },
        right: { style: 'thin' }
      };
    });
    
    allCustomers.forEach(customer => {
      const row = newCustomerSheet.addRow({
        'Customer Mobile Number': customer['Customer Mobile Number'] || '',
        'Flat Number': customer['Flat Number'] || ''
      });
      
      row.eachCell(cell => {
        cell.font = { size: 12 };
        cell.alignment = { horizontal: 'left' };
        cell.border = {
          top: { style: 'thin' },
          bottom: { style: 'thin' },
          left: { style: 'thin' },
          right: { style: 'thin' }
        };
      });
    });
  }
}