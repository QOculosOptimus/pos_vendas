function doGet() {
    return HtmlService.createHtmlOutputFromFile('index');
}

function getRedirectUrl() {
    const clientId = '7d7db940a604900abdaf3641fe304423fac2d65c';
    const redirectUri = encodeURIComponent('https://developer.bling.com.br/oauth/redirect');
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

