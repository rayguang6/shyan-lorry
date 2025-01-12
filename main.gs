/*****************************************************
 *  onOpen()
 *  Adds a custom menu for easy access.
 *****************************************************/
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Custom Tools')
    .addItem('Print Lorry', 'showFlexLayoutForActiveSheetWithSums')
    .addItem('Print Customer Total', 'generateSimpleTotalsLayout')
    .addToUi();
}

/*****************************************************
 *  getSheetMetadata()
 *  Fetches lorry details (keys and values from columns A and B).
 *****************************************************/
function getSheetMetadata() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  const metadata = [];

  // Loop through rows 1-3 and fetch both key and value
  for (let row = 1; row <= 3; row++) {
    const key = sheet.getRange(`A${row}`).getValue();
    const value = sheet.getRange(`B${row}`).getValue();
    if (key && value) {
      metadata.push({ key, value });
    }
  }

  return metadata;
}

/*****************************************************
 *  getCustomerDataFromActiveSheet()
 *  Reads customer data starting from row 5.
 *****************************************************/
function getCustomerDataFromActiveSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const activeSheet = ss.getActiveSheet();

  // Get all data starting from row 5
  const lastColumn = activeSheet.getLastColumn();
  const values = activeSheet.getRange(5, 1, activeSheet.getLastRow() - 4, lastColumn).getValues();
  if (!values || values.length < 2) {
    return []; // No data or only a header row
  }

  const header = values[0]; // Customer names
  const dataRows = values.slice(1);

  const result = []; // Initialize as an array

  // Loop through each column to process customer data
  for (let col = 1; col < header.length; col++) {
    const customerName = header[col];
    const itemsForThisCustomer = [];

    for (let r = 0; r < dataRows.length; r++) {
      const itemName = dataRows[r][0]; // Item name in column A
      const quantity = Number(dataRows[r][col] || 0);
      if (quantity > 0) {
        itemsForThisCustomer.push({ item: itemName, qty: quantity });
      }
    }

    if (itemsForThisCustomer.length > 0) {
      const totalQty = itemsForThisCustomer.reduce((sum, obj) => sum + obj.qty, 0);
      result.push({ customerName, items: itemsForThisCustomer, totalQty });
    }
  }

  // Uncomment the following line to enable sorting by totalQty in descending order
  result.sort((a, b) => b.items.length - a.items.length);

  return result; // Return an array of customer data
}

/*****************************************************
 *  buildFlexboxHtmlWithSums(customerData, sheetName, metadata)
 *  Builds the HTML string for the modal dialog.
 *****************************************************/
function buildFlexboxHtmlWithSums(customerData, sheetName, metadata) {
  const css = `
    <style>
      body {
        margin: 0;
        padding: 0;
        font-family: Arial, sans-serif;
        background-color: #F7F7F7;
      }
      .button-container {
        text-align: center;
        margin-bottom: 10px;
      }
      .lorry-details {
        display: flex;
        justify-content: center;
        align-items: center;
        gap: 20px;
        margin: 10px 0;
        font-size: 14px;
        font-weight: bold;
        page-break-after: avoid;
      }
      .lorry-details div {
        white-space: nowrap; /* Prevent line breaks */
      }
      .page-container {
        max-width: 210mm;
        margin: 0 auto;
        padding: 10px;
        background-color: #FFFFFF;
        box-shadow: 0 0 10px rgba(0, 0, 0, 0.1);
      }
      h1 {
        text-align: center;
        margin-bottom: 10px;
        page-break-after: avoid;
      }
      .container {
        display: flex;
        flex-wrap: wrap;
        gap: 15px;
        justify-content: center;
        page-break-inside: avoid;
      }
      .box {
        width: 180px;
        border: 1px solid #CCC;
        padding: 10px;
        box-sizing: border-box;
        background-color: #FAFAFA;
        page-break-inside: avoid;
        display: flex;
        flex-direction: column;
      }
      .box h2 {
        margin: 0 0 10px;
        font-size: 16px;
        text-align: center;
      }
      table {
        width: 100%;
        border-collapse: collapse;
      }
      table tr {
        border-bottom: 1px solid #EEE;
      }
      td {
        padding: 5px 0;
        font-size: 14px;
      }
      td.qty {
        text-align: right;
      }
      .total-container {
        margin-top: auto;
        margin-bottom: 0;
        display: flex;
        align-items: center;
        justify-content: space-between;
        font-weight: bold;
      }
      .open-tab-btn {
        display: inline-block;
        padding: 8px 16px;
        background: #4285f4;
        color: #fff;
        text-decoration: none;
        border-radius: 4px;
        cursor: pointer;
        font-size: 14px;
        border: none;
      }
      .open-tab-btn:hover {
        background: #3367d6;
      }
      @media print {
        .button-container {
          display: none;
        }
        .lorry-details {
          page-break-after: avoid;
        }
        h1 {
          page-break-after: avoid;
        }
        .container {
          page-break-inside: avoid;
        }
        .box {
          page-break-inside: avoid;
        }
      }
    </style>
  `;

  let html = `<!DOCTYPE html><html><head><meta charset="utf-8">${css}</head><body>`;
  
  // Add the "Open in New Tab" button
  html += `
    <div class="button-container">
      <button class="open-tab-btn" onclick="openInNewTab()">Open in New Tab</button>
    </div>
  `;

  // Add the title
  html += `
    <div class="page-container">
      <h1>${sheetName}</h1>
  `;

  // Add the lorry details in a single row
  html += `
      <div class="lorry-details">
  `;
  metadata.forEach(({ key, value }) => {
    html += `<div>${key}: ${value}</div>`;
  });
  html += `</div>`;

  // Add the customer data
  html += `<div class="container">`;

  customerData.forEach(({ customerName, items, totalQty }) => {
    html += `
      <div class="box">
        <h2>${customerName}</h2>
        <table>
    `;

    items.forEach(({ item, qty }) => {
      html += `
        <tr>
          <td>${item}</td>
          <td class="qty">${qty}</td>
        </tr>
      `;
    });

    html += `
        </table>
        <div class="total-container">
          <h5>TOTAL:</h5>
          <p>${totalQty}</p>
        </div>
      </div>
    `;
  });

  html += `
      </div>
    </div>
    <script>
      function openInNewTab() {
        var newWin = window.open('', '_blank');
        newWin.document.write(document.documentElement.innerHTML);
        newWin.document.close();
        newWin.focus();
      }
    </script>
  </body></html>`;
  return html;
}


 /*****************************************************
 *  calculateCustomerTotals()
 *  Reads customer data and calculates total quantities.
 *****************************************************/
function calculateCustomerTotals() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const dataRange = sheet.getDataRange();
  const data = dataRange.getValues();

  if (data.length < 2) {
    return []; // No data or just headers
  }

  const customerHeaders = data[0]; // First row: customer names
  const quantityRows = data.slice(1); // Rows below headers: quantities

  const totals = customerHeaders.map((customerName, colIndex) => {
    // if (colIndex === 0) return null; // Skip first column if it's not a customer
    const totalQuantity = quantityRows.reduce(
      (sum, row) => sum + (Number(row[colIndex]) || 0),
      0
    );
    return { customerName, totalQuantity };
  }).filter(Boolean); // Remove nulls from non-customer columns

  return totals;
}

/*****************************************************
 *  buildSimpleHtmlLayout()
 *  Generates an HTML table layout with fixed alternate row colors for printing.
 *****************************************************/
function buildSimpleHtmlLayout(customerTotals, title) {
  const styles = `
    <style>
      body {
        font-family: Arial, sans-serif;
        background: #f7f7f7;
        margin: 0;
        padding: 0;
        display: flex;
        justify-content: center;
        align-items: center;
        height: 100vh;
        flex-direction: column;
      }
      .table-container {
        background: white;
        border-radius: 8px;
        box-shadow: 0 4px 8px rgba(0, 0, 0, 0.1);
        overflow: auto;
        width: 100%;
        max-width: 600px;
        max-height: 90vh;
        padding: 10px;
      }
      h1 {
        font-size: 18px;
        font-weight: bold;
        text-align: center;
        margin: 10px 0;
      }
      table {
        width: 100%;
        border-collapse: collapse;
        margin-top: 10px;
      }
      tr {
        border: 1px solid #ddd; /* Add borders to rows */
      }
      tr.row-even {
        background: #f2f2f2; /* Alternate row color */
      }
      tr:hover {
        background: #e8f0fe; /* Highlight on hover */
      }
      td {
        font-size: 14px;
        padding: 10px;
        text-align: left;
        border-left: 1px solid #ddd;
        border-right: 1px solid #ddd;
      }
      td:nth-child(2) {
        text-align: right; /* Align totals to the right */
      }
      .total-row {
        font-weight: bold;
        color: black;
      }
      .button-container {
        width: 100%;
        text-align: center;
        margin-bottom: 10px;
      }
      .button {
        padding: 10px 20px;
        font-size: 14px;
        background: #4285f4;
        color: white;
        border: none;
        border-radius: 5px;
        cursor: pointer;
        text-transform: uppercase;
        font-weight: bold;
      }
      .button:hover {
        background: #3367d6;
      }
      @media print {
        .button-container {
          display: none; /* Hide the button when printing */
        }
        body {
          margin: 0;
          padding: 0;
          height: auto;
        }
        .table-container {
          overflow: visible;
          max-height: unset;
        }
        tr.row-even {
          background: #f2f2f2 !important; /* Force alternate colors for printing */
        }
        tr {
          border: 1px solid #ddd; /* Retain borders for printed rows */
        }
      }
    </style>
  `;

  // Calculate the grand total
  const grandTotal = customerTotals.reduce((sum, { totalQuantity }) => sum + totalQuantity, 0);

  let html = `
    <!DOCTYPE html>
    <html>
    <head>
      <title>${title}</title>
      ${styles}
    </head>
    <body>
      <div class="button-container">
        <button class="button" onclick="openInNewTab()">Open in New Tab</button>
      </div>
      <div class="table-container">
        <h1>${title}</h1>
        <table>
          <tbody>
            ${customerTotals
              .map(
                ({ customerName, totalQuantity }, index) => `
              <tr class="${index % 2 === 0 ? 'row-even' : ''}">
                <td>${customerName}</td>
                <td>${totalQuantity}</td>
              </tr>
            `
              )
              .join('')}
            <tr class="total-row">
              <td>Total</td>
              <td>${grandTotal}</td>
            </tr>
          </tbody>
        </table>
      </div>
      <script>
        function openInNewTab() {
          const newTab = window.open('', '_blank');
          newTab.document.write(document.documentElement.outerHTML);
          newTab.document.close();
        }
      </script>
    </body>
    </html>
  `;

  return html;
}


/*****************************************************
 *  generateSimpleTotalsLayout()
 *  Displays the layout with enhanced row visibility and title in a modal dialog.
 *****************************************************/
function generateSimpleTotalsLayout() {
  const customerTotals = calculateCustomerTotals();
  const sheetTitle = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getName(); // Use sheet name as title
  const html = buildSimpleHtmlLayout(customerTotals, `${sheetTitle}`);

  const htmlOutput = HtmlService.createHtmlOutput(html)
    .setWidth(600)
    .setHeight(800); // Adjusted size for better visibility

  SpreadsheetApp.getUi().showModalDialog(htmlOutput, ' ');
}



/*****************************************************
 *  showFlexLayoutForActiveSheetWithSums()
 *  Displays the modal with the generated HTML.
 *****************************************************/
function showFlexLayoutForActiveSheetWithSums() {
  const metadata = getSheetMetadata();
  const customerData = getCustomerDataFromActiveSheet(); // Ensure the array is passed here
  const sheetName = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getName();
  const htmlString = buildFlexboxHtmlWithSums(customerData, sheetName, metadata);

  const htmlOutput = HtmlService.createHtmlOutput(htmlString).setWidth(1200).setHeight(1000);
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Print Layout');
}
