const sheetId = "1b2ReDDV_cPomVDR0sPcoFuQEhd8brGBnF7AFblQTIvY";
const authSheetName = "AuthStuff";
const vendasSheetName = "vendas";
const produtosSheetName = "produtos";

function doGet(e) {
    if (e.parameter.code && e.parameter.state) {
        const sheet = SpreadsheetApp.openById(sheetId).getSheetByName(authSheetName);
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
    const sheet = SpreadsheetApp.openById(sheetId).getSheetByName(authSheetName);
    const expirationTime = sheet.getRange("B7").getValue();
    const currentTime = new Date();
    return currentTime >= expirationTime;
}

function fetchSalesOrders() {
    // Open the spreadsheet and access the AuthStuff sheet to get the access token
    const spreadsheet = SpreadsheetApp.openById(sheetId);
    const authSheet = spreadsheet.getSheetByName(authSheetName);
    const accessToken = authSheet.getRange("B6").getValue();

    // Define the API endpoint URL
    const url = 'https://developer.bling.com.br/api/bling/pedidos/vendas?pagina=1&limite=100&dataInicial=2024-06-01&dataFinal=2024-12-15';

    // Fetch the data from the API
    const response = UrlFetchApp.fetch(url, {
        method: 'GET',
        headers: {
            'Authorization': `Bearer ${accessToken}`,
            'accept': 'application/json'
        }
    });

    // Parse the JSON response
    const jsonResponse = JSON.parse(response.getContentText());

    // Access the 'vendas' sheet
    const vendasSheet = spreadsheet.getSheetByName(vendasSheetName);

    // Define the header row based on the specified columns
    const headers = [
        "ID",
        "Número",
        "Data",
        "Data Saída",
        "Data Prevista",
        "Total Produtos",
        "Total",
        "Situação Valor",
        "Situação ID",
        "Número Loja",
        "Loja ID",
        "Contato Tipo Pessoa",
        "Contato Nome",
        "Contato ID",
        "Contato Documento"
    ];

    // Clear existing content in the 'vendas' sheet
    vendasSheet.clearContents();

    // Set the header row
    vendasSheet.getRange(1, 1, 1, headers.length).setValues([headers]);

    // Prepare the data rows
    const dataRows = jsonResponse.data.map(order => [
        order.id,
        order.numero,
        order.data,
        order.dataSaida,
        order.dataPrevista,
        order.totalProdutos,
        order.total,
        order.situacao.valor,
        order.situacao.id,
        order.numeroLoja,
        order.loja.id,
        order.contato.tipoPessoa,
        order.contato.nome,
        order.contato.id,
        order.contato.numeroDocumento
    ]);

    // Check if there are data rows to insert
    if (dataRows.length > 0) {
        // Determine the range to insert data
        vendasSheet.getRange(2, 1, dataRows.length, headers.length).setValues(dataRows);
    }

    // Return the parsed JSON response
    return jsonResponse;
}

function fetchProducts() {
    // Open the spreadsheet and access the AuthStuff sheet to get the access token
    const spreadsheet = SpreadsheetApp.openById(sheetId);
    const authSheet = spreadsheet.getSheetByName(authSheetName);
    const accessToken = authSheet.getRange("B6").getValue();

    // Define the API endpoint URL
    const url = 'https://developer.bling.com.br/api/bling/produtos?pagina=1&limite=100&criterio=1&tipo=T&dataInclusaoInicial=2024-06-01%2012%3A00%3A00&dataInclusaoFinal=2024-12-01%2013%3A00%3A00&nome=%20';

    // Fetch the data from the API
    const response = UrlFetchApp.fetch(url, {
        method: 'GET',
        headers: {
            'Authorization': `Bearer ${accessToken}`,
            'accept': 'application/json'
        }
    });

    // Parse the JSON response
    const jsonResponse = JSON.parse(response.getContentText());

    // Access the 'produtos' sheet
    const produtosSheet = spreadsheet.getSheetByName(produtosSheetName);

    // Define the header row based on the specified columns
    const headers = [
        "ID",
        "Nome",
        "Código",
        "Preço",
        "Preço de Custo",
        "Saldo Virtual Total",
        "Tipo",
        "Situação",
        "Formato",
        "Descrição Curta",
        "Imagem URL"
    ];

    // Clear existing content in the 'produtos' sheet
    produtosSheet.clearContents();

    // Set the header row
    produtosSheet.getRange(1, 1, 1, headers.length).setValues([headers]);

    // Prepare the data rows
    const dataRows = jsonResponse.data.map(product => [
        product.id,
        product.nome,
        product.codigo,
        product.preco,
        product.precoCusto,
        product.estoque.saldoVirtualTotal,
        product.tipo,
        product.situacao,
        product.formato,
        product.descricaoCurta,
        product.imagemURL
    ]);

    // Check if there are data rows to insert
    if (dataRows.length > 0) {
        // Determine the range to insert data
        produtosSheet.getRange(2, 1, dataRows.length, headers.length).setValues(dataRows);
    }

    // Return the parsed JSON response
    return jsonResponse;
}
