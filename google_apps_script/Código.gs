function doGet(e) {
    const path = e.pathInfo || "";
    
    // Check if the request is for "/callback" or if it has "code" and "state" parameters
    if (path.startsWith("callback")) {
        return handleCallback(e);
    } else if (e.parameter.code && e.parameter.state) {
        const sheetId = "1b2ReDDV_cPomVDR0sPcoFuQEhd8brGBnF7AFblQTIvY";
        const sheet = SpreadsheetApp.openById(sheetId).getSheetByName("Página1");
        sheet.getRange("D1").setValue(e.parameter.code);
        sheet.getRange("D2").setValue(e.parameter.state);

        const allParams = Object.keys(e.parameter).map(key => `${key}: ${e.parameter[key]}`).join(", ");
        sheet.getRange("D3").setValue(allParams);
        return HtmlService.createHtmlOutputFromFile('index');
    }
    else {
        const sheetId = "1b2ReDDV_cPomVDR0sPcoFuQEhd8brGBnF7AFblQTIvY";
        const sheet = SpreadsheetApp.openById(sheetId).getSheetByName("Página1");
        sheet.getRange("C1").setValue(path);
        return HtmlService.createHtmlOutputFromFile('index');
    }
}


function getRedirectUrl() {
    const clientId = '7d7db940a604900abdaf3641fe304423fac2d65c';
    // const redirectUri = encodeURIComponent('https://developer.bling.com.br/oauth/redirect');
    const redirectUri = encodeURIComponent('https://script.google.com/macros/s/AKfycbxnGdqTfhojDWvNS_0igaq6ht8fgnL5G_sBmr8Dwo8/dev');
    const responseType = 'code';
    const state = encodeURIComponent(generateRandomString(70) + '==');
    const url = `https://www.bling.com.br/OAuth2/views/login.php?client_id=${clientId}&redirect_uri=${redirectUri}&response_type=${responseType}&state=${state}`;
    return url;
}

// Helper function to generate a random string for the state parameter
function generateRandomString(length) {
    const characters = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789';
    let result = '';
    for (let i = 0; i < length; i++) {
        result += characters.charAt(Math.floor(Math.random() * characters.length));
    }
    return result;
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


