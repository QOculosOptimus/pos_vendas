const sheetId = "1b2ReDDV_cPomVDR0sPcoFuQEhd8brGBnF7AFblQTIvY";
const sheetName = "Página1";

function doGet(e) {
    if (e.parameter.code && e.parameter.state) {
        const sheet = SpreadsheetApp.openById(sheetId).getSheetByName(sheetName);
        sheet.getRange("B1").setValue(e.parameter.code);
        sheet.getRange("B2").setValue(new Date());
        sheet.getRange("B3").setValue(e.parameter.state);
        sheet.getRange("B4").setValue(new Date());

        // Exchange the authorization code for a bearer token
        const tokenData = fetchBearerToken(e.parameter.code);
        if (tokenData) {
            sheet.getRange("B5").setValue(tokenData.refresh_token);
            sheet.getRange("B6").setValue(tokenData.access_token);
            
            // Calculate the expiration time and record it in C7
            const expirationTime = new Date();
            expirationTime.setSeconds(expirationTime.getSeconds() + tokenData.expires_in);
            sheet.getRange("B7").setValue(expirationTime);
        }
    }
    return HtmlService.createHtmlOutputFromFile('index');
}

function fetchBearerToken(authCode) {
    const clientId = '7d7db940a604900abdaf3641fe304423fac2d65c';
    const clientSecret = '39c1816d51a67dfc30a1eb1d8fa7b8341442fdc6c00104c635b37bbb93bb';
    const redirectUri = 'https://script.google.com/macros/s/AKfycbxnGdqTfhojDWvNS_0igaq6ht8fgnL5G_sBmr8Dwo8/dev';
    
    const tokenUrl = 'https://developer.bling.com.br/api/bling/oauth/token';
    const payload = {
        method: 'post',
        payload: {
            grant_type: 'authorization_code',
            code: authCode,
            redirect_uri: redirectUri,
            client_id: clientId,
            client_secret: clientSecret,
        },
    };

    const response = UrlFetchApp.fetch(tokenUrl, payload);
    const responseData = JSON.parse(response.getContentText());
    
    // Return token data including access token, refresh token, and expires_in
    return {
        access_token: responseData.access_token || null,
        refresh_token: responseData.refresh_token || null,
        expires_in: responseData.expires_in || null,
    };
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

function checkTokenExpiration() {
    const sheet = SpreadsheetApp.openById(sheetId).getSheetByName(sheetName);
    const expirationTime = sheet.getRange("B7").getValue();
    const currentTime = new Date();
    return currentTime >= expirationTime;
}

function fetchSalesOrders() {
    const sheet = SpreadsheetApp.openById(sheetId).getSheetByName(sheetName);
    const accessToken = sheet.getRange("B6").getValue();
    const url = 'https://developer.bling.com.br/api/bling/pedidos/vendas?pagina=1&limite=100&dataInicial=2024-06-01&dataFinal=2024-12-15';
    
    const response = UrlFetchApp.fetch(url, {
        method: 'GET',
        headers: {
            'Authorization': `Bearer ${accessToken}`,
            'accept': 'application/json'
        }
    });
    
    return JSON.parse(response.getContentText());
}

