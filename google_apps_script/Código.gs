const sheetId = "1b2ReDDV_cPomVDR0sPcoFuQEhd8brGBnF7AFblQTIvY";
const authSheetName = "AuthStuff";
const vendasSheetName = "vendas";
const produtosSheetName = "produtos";
const vendasItensSheetName = "vendas_itens";

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

    // Access the 'vendas' sheet
    const vendasSheet = spreadsheet.getSheetByName(vendasSheetName);

    // Get the last order date from the sheet, or use 2024-06-01 as the start date
    const lastOrderDate = getLastOrderDate(vendasSheet) || '2024-06-01';

    // Define the API base URL and parameters
    const baseUrl = 'https://developer.bling.com.br/api/bling/pedidos/vendas';
    const endDate = new Date(); // Today
    const formattedEndDate = formatDate(endDate);

    // Initialize pagination
    let page = 1;
    let allOrders = [];

    // Loop to fetch data page by page until all orders are retrieved
    while (true) {
	console.log('Fetching page', page);
        const url = `${baseUrl}?pagina=${page}&limite=100&dataInicial=${lastOrderDate}&dataFinal=${formattedEndDate}`;

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

        // If no data is returned, break out of the loop
        if (!jsonResponse.data || jsonResponse.data.length === 0) {
            break;
        }

        // Append the data to the allOrders array
        allOrders = allOrders.concat(jsonResponse.data);

        // Check if we have fewer than 100 orders, indicating it's the last page, if it's not the last page, wait 5 seconds to avoid rate limiting
        if (jsonResponse.data.length < 100) {
            break;
        } else {
	    Utilities.sleep(5000);
	}

        // Move to the next page
        page++;
    }

    // Invert the order of "allOrders"
    allOrders.reverse();

    // Prepare the data rows
    const dataRows = allOrders.map(order => [
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
    vendasSheet.getRange(1, 1, 1, headers.length).setValues([headers]);

    // Set the data rows starting from the last row + 1 of the sheet or the first row that has the third column equal to the last order date
    if (dataRows.length > 0) {
        const lastRow = vendasSheet.getLastRow();
	const lastDateRow = vendasSheet.getRange("C:C").createTextFinder(lastOrderDate).findNext();
	console.log('lastRow:', lastRow);
	// console.log('lastDateRow:', lastDateRow);
	const startRow = lastDateRow ? lastDateRow.getRow() : lastRow + 1;
	console.log('startRow:', startRow);
	vendasSheet.getRange(startRow, 1, dataRows.length, headers.length).setValues(dataRows);
    }

    return true;
}

// Function to get the last order date from the 'vendas' sheet
function getLastOrderDate(vendasSheet) {
    const lastRow = vendasSheet.getLastRow();
    if (lastRow < 2) return null; // No data in the sheet

    const lastDate = vendasSheet.getRange(lastRow, 3).getValue(); // Assuming 'Data' is in the 3rd column
    // Convert the string to a Date object
    var dateObj = new Date(lastDate);
    
    // Format the date as YYYY-MM-DD
    var formattedDate = Utilities.formatDate(dateObj, Session.getScriptTimeZone(), "yyyy-MM-dd");
    
    console.log(formattedDate); // Outputs: 2024-11-15
    return formattedDate || null;
}

// Helper function to format dates as yyyy-mm-dd
function formatDate(date) {
    const year = date.getFullYear();
    const month = String(date.getMonth() + 1).padStart(2, '0'); // Months are 0-indexed
    const day = String(date.getDate()).padStart(2, '0');
    return `${year}-${month}-${day}`;
}

function fetchItemsOrders() {
    // Open the spreadsheet and access the AuthStuff sheet to get the access token
    const spreadsheet = SpreadsheetApp.openById(sheetId);
    const authSheet = spreadsheet.getSheetByName(authSheetName);
    const accessToken = authSheet.getRange("B6").getValue();

    // Access the 'vendas' and 'vendas_itens' sheets
    const vendasSheet = spreadsheet.getSheetByName(vendasSheetName);
    const vendasItensSheet = spreadsheet.getSheetByName(vendasItensSheetName);

    // Get the list of order IDs from 'vendas' sheet
    const vendasData = vendasSheet.getDataRange().getValues(); // Includes headers
    if (vendasData.length < 2) {
        console.log('No orders found in vendas sheet');
        return;
    }
    const vendasHeaders = vendasData[0];
    const vendasRows = vendasData.slice(1); // Exclude headers
    const vendasIdIndex = vendasHeaders.indexOf('ID'); // Column index for order ID
    if (vendasIdIndex === -1) {
        console.log('ID column not found in vendas sheet');
        return;
    }
    const vendasOrderIds = vendasRows.map(row => row[vendasIdIndex]);

    // Get processed order IDs from 'vendas_itens' sheet
    const vendasItensData = vendasItensSheet.getDataRange().getValues();
    let processedOrderIdsSet = new Set();
    if (vendasItensData.length > 1) {
        const vendasItensHeaders = vendasItensData[0];
        const vendasItensRows = vendasItensData.slice(1);
        const vendasItensOrderIdIndex = vendasItensHeaders.indexOf('Order ID'); // Column index for order ID
        if (vendasItensOrderIdIndex === -1) {
            console.log('Order ID column not found in vendas_itens sheet');
            return;
        }
        vendasItensRows.forEach(row => processedOrderIdsSet.add(row[vendasItensOrderIdIndex]));
    }

    // Identify unprocessed orders
    const unprocessedOrderIds = vendasOrderIds.filter(id => !processedOrderIdsSet.has(id));

    console.log(`Found ${unprocessedOrderIds.length} unprocessed orders`);

    // Determine the next available row in 'vendas_itens' sheet
    let lastRow = vendasItensSheet.getLastRow();

    // Write headers if sheet is empty
    if (lastRow < 1) {
        const headers = [
            "Order ID",
            "Item ID",
            "Código",
            "Unidade",
            "Quantidade",
            "Desconto",
            "Valor",
            "Aliquota IPI",
            "Descrição",
            "Produto ID",
            "Comissão Base",
            "Comissão Aliquota",
            "Comissão Valor"
        ];
        vendasItensSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
        lastRow = 1;
    }

    // Set the starting row
    let nextRow = lastRow + 1;

    for (const orderId of unprocessedOrderIds) {
        const url = `https://developer.bling.com.br/api/bling/pedidos/vendas/${orderId}`;

        try {
            const response = UrlFetchApp.fetch(url, {
                method: 'GET',
                headers: {
                    'Authorization': `Bearer ${accessToken}`,
                    'accept': 'application/json'
                }
            });

            const jsonResponse = JSON.parse(response.getContentText());

            if (!jsonResponse.data) {
                console.log(`No data for order ID ${orderId}`);
                continue;
            }

            const orderData = jsonResponse.data;
            const orderItems = orderData.itens || [];

            const dataRows = [];

            for (const item of orderItems) {
                const dataRow = [
                    orderData.id,               // Order ID
                    item.id,                    // Item ID
                    item.codigo,                // Código
                    item.unidade,               // Unidade
                    item.quantidade,            // Quantidade
                    item.desconto,              // Desconto
                    item.valor,                 // Valor
                    item.aliquotaIPI,           // Aliquota IPI
                    item.descricao,             // Descrição
                    item.produto.id,            // Produto ID
                    item.comissao.base,         // Comissão Base
                    item.comissao.aliquota,     // Comissão Aliquota
                    item.comissao.valor         // Comissão Valor
                ];

                dataRows.push(dataRow);
            }

            if (dataRows.length > 0) {
                vendasItensSheet.getRange(nextRow, 1, dataRows.length, dataRows[0].length).setValues(dataRows);
                console.log(`Wrote ${dataRows.length} items for order ID ${orderId} to vendas_itens sheet`);
                nextRow += dataRows.length;
            }

            // Add orderId to processedOrderIdsSet
            processedOrderIdsSet.add(orderId);

            // Wait to avoid rate limits
            Utilities.sleep(1000);

        } catch (error) {
            console.log(`Error fetching order ID ${orderId}: ${error}`);
            continue;
        }
    }
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

