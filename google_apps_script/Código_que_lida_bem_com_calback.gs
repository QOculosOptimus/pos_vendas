function doGet(e) {
    const path = e.pathInfo || "";
    
    // Check if the request is specifically for "/callback"
    if (path.startsWith("callback")) {
        return handleCallback(e);
    } else {
        // return handleNonCallback(e);
        return HtmlService.createHtmlOutputFromFile('index');
    }
}

function handleNonCallback(e){
  const pathInfo = e.pathInfo || "";

  const sheetId = "1b2ReDDV_cPomVDR0sPcoFuQEhd8brGBnF7AFblQTIvY";
  const sheet = SpreadsheetApp.openById(sheetId).getSheetByName("Página1");

  // Put the values in the specified cells
  sheet.getRange("C1").setValue(pathInfo);

  // Respond to the callback request
  return ContentService.createTextOutput("Callback processed successfully (handleNonCallback).");
}


function handleCallback(e) {
    const queryString = e.pathInfo;
    const params = parseQueryString(queryString);
    const code = params.code || queryString;
    const state = params.state || "veio nada no state";

    const sheetId = "1b2ReDDV_cPomVDR0sPcoFuQEhd8brGBnF7AFblQTIvY";
    const sheet = SpreadsheetApp.openById(sheetId).getSheetByName("Página1");

    // Put the values in the specified cells
    sheet.getRange("B1").setValue(code);
    sheet.getRange("B3").setValue(state);

    // Respond to the callback request
    return ContentService.createTextOutput("Callback processed successfully (handleCallback).");
}

function parseQueryString(queryString) {
    const params = {};
    const pairs = (queryString || "").split("&");
    pairs.forEach(pair => {
        const [key, value] = pair.split("=");
        if (key && value) {
            params[decodeURIComponent(key)] = decodeURIComponent(value);
        }
    });
    return params;
}
