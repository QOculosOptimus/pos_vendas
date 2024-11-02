const sheetId = "1b2ReDDV_cPomVDR0sPcoFuQEhd8brGBnF7AFblQTIvY";
const sheetName = "Página1";

function doGet(e) {
    const path = e.pathInfo || "";
    
    if (e.parameter.code && e.parameter.state) {
        const sheet = SpreadsheetApp.openById(sheetId).getSheetByName(sheetName);
        sheet.getRange("B1").setValue(e.parameter.code);
        sheet.getRange("B2").setValue(new Date());
        sheet.getRange("B3").setValue(e.parameter.state);
	sheet.getRange("B4").setValue(new Date());
    }
    return HtmlService.createHtmlOutputFromFile('index');
}


function getRedirectUrl() {
    const clientId = '7d7db940a604900abdaf3641fe304423fac2d65c';
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
