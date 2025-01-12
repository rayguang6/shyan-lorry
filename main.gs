/*****************************************************
 *  onOpen()
 *  Adds a custom menu for easy access.
 *****************************************************/
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Custom Tools')
    .addItem('Generate Print Layout (with Sums)', 'showFlexLayoutForActiveSheetWithSums')
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
  // const values = activeSheet.getRange('A5:Z').getValues();
  if (!values || values.length < 2) {
    return {}; // No data or only a header row
  }

  const header = values[0]; // Customer names
  const dataRows = values.slice(1);

  const result = {};

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
      result[customerName] = itemsForThisCustomer;
    }
  }

  return result;
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
        display:flex;
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
      .sumRow {
        font-weight: bold;
        text-align: right;
        padding-top: 8px;
      }
      .sumLabel {
        padding-right: 5px;
        color: #555;
      }
      .total-container{
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

  const customerNames = Object.keys(customerData);
  if (customerNames.length > 0) {
    customerNames.forEach((customerName) => {
      const items = customerData[customerName];
      const sum = items.reduce((acc, obj) => acc + obj.qty, 0);

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
            <p>${sum}</p>
          </div>
        </div>
      `;
    });
  }

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
 *  showFlexLayoutForActiveSheetWithSums()
 *  Displays the modal with the generated HTML.
 *****************************************************/
function showFlexLayoutForActiveSheetWithSums() {
  const metadata = getSheetMetadata();
  const data = getCustomerDataFromActiveSheet();
  const sheetName = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getName();
  const htmlString = buildFlexboxHtmlWithSums(data, sheetName, metadata);

  const htmlOutput = HtmlService.createHtmlOutput(htmlString).setWidth(1200).setHeight(1000);
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Print Layout');
}
