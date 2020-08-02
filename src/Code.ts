function reset() {
    const service = getService();
    service.reset();
}

function getService() {
    // @ts-ignore
    return OAuth2.createService('questrade')
        .setAuthorizationBaseUrl('https://login.questrade.com/oauth2/authorize')
        .setTokenUrl('https://login.questrade.com/oauth2/token')
        .setClientId(PropertiesService.getScriptProperties().getProperty('consumerKey'))
        .setClientSecret('secret') // No secret provided by QT. Use dummy one to make oauth2 lib happy.
        .setCallbackFunction('authCallback')
        .setPropertyStore(PropertiesService.getUserProperties())
        .setScope('read_acc')
        .setParam('response_type', 'code');
}

function authCallback(request) {
    const service = getService();
    const authorized = service.handleCallback(request);
    if (authorized) {
        return HtmlService.createHtmlOutput('Success! <script>setTimeout(function() { top.window.close() }, 1);</script>');
    } else {
        return HtmlService.createHtmlOutput('Denied.');
    }
}

function logRedirectUri() {
    // @ts-ignore
    console.log(OAuth2.getRedirectUri());
}

const QuestradeApiSession = function () {
    const service = getService();
    if (!service.hasAccess()) {
        const authorizationUrl = service.getAuthorizationUrl();
        const template: GoogleAppsScript.HTML.HtmlTemplate = HtmlService.createTemplate(
            '<a href="<?= authorizationUrl ?>" target="_blank">Authorize</a>. ' +
            'Pull again when the authorization is complete.');
        // @ts-ignore
        template.authorizationUrl = authorizationUrl;
        const page = template.evaluate();
        SpreadsheetApp.getUi().showSidebar(page);
        return;
    }

    const apiOptions: GoogleAppsScript.URL_Fetch.URLFetchRequestOptions = {
        'method': 'get',
        'headers': {
            'Authorization': 'Bearer ' + service.getAccessToken()
        }
    };
    const apiUrl = service.getToken().api_server;

    const fetch = function(query) {
        return JSON.parse(UrlFetchApp.fetch(apiUrl + query, apiOptions).getContentText());
    }

    const getAccounts = this.getAccounts = function() {
        return fetch('v1/accounts').accounts;
    };
    this.accounts = this.getAccounts();

    const getPositions = this.getPositions = function(accountNumber) {
        return fetch('v1/accounts/' + accountNumber + '/positions').positions;
    };

    const getSymbols = this.getSymbols = function(symbolIds) {
        return fetch('v1/symbols?ids=' + symbolIds.join(',')).symbols;
    };

    const getBalances = this.getBalances = function(accountNumber) {
        return fetch('v1/accounts/' + accountNumber + '/balances').perCurrencyBalances;
    }

    const getExchangeRates = (base) => {
        return JSON.parse(UrlFetchApp.fetch(`https://api.exchangeratesapi.io/latest?base=${base}`).getContentText()).rates;
    };

    const prefixKeys = (map, prefix) => {
        if (map) {
            return Object.entries(map).reduce((m, [k,v]) => { m[prefix + k] = v; return m; }, {});
        }
    };

    const getEnrichedPositions = this.getEnrichedPositions = function() {
        const positions = getAccounts().flatMap(ac => {
            const ac2 = prefixKeys(ac, 'account.');
            const acPositions = getPositions(ac.number).map(pos => ({
                ...pos, ...ac2
            }));

            getBalances(ac.number).forEach(bal => {
                acPositions.push({
                    symbol: '$' + bal.currency,
                    symbolId: '$' + bal.currency,
                    currentMarketValue: bal.cash,
                    ...ac2
                })
            });

            return acPositions;
        });

        const rates = getExchangeRates('CAD');
        const getValueInCAD = (value, currency) => value / rates[currency];

        const symbols = [ ...getSymbols(positions.map(pos => pos.symbolId).filter(id => Number.isInteger(id))),
            { symbolId: '$CAD', currency: 'CAD' }, { symbolId: '$USD', currency: 'USD' } ];

        return positions.map(pos => {
            const symbol = symbols.find(s => s.symbolId == pos.symbolId);
            const exchangeRate = 1.0 / rates[symbol.currency];
            const valueInCAD = pos.currentMarketValue * exchangeRate;
            return { 
                ...pos,
                ...prefixKeys(symbol, 'symbol.'),
                exchangeRate: exchangeRate,
                valueInCAD: valueInCAD,
                currency: symbol.currency
            };
        });
    };
}

function updatePositions() {
    const qt = new QuestradeApiSession();
    const positions = qt.getEnrichedPositions();
    // console.log('positions', positions);

    const doc = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = doc.getSheetByName('Positions');
    const values = sheet.getDataRange().getValues();
    var lastRowIndex = values.length;

    const headerRow = values[0];
    const numHeaderRows = 2; // header + totals
    const columnMap = headerRow.reduce((map, val, index) => {
        // @ts-ignore
        map[val] = index;
        return map;
    }, {});

    const accountNumCol = columnMap['account.number'];
    const symbolIdCol = columnMap['symbolId'];
    const updatedRowIndicies = new Set();

    positions.forEach(pos => {
        var rowIndex = values.findIndex(row => row[accountNumCol] == pos['account.number'] && row[symbolIdCol] == pos.symbolId);
        if (rowIndex < 0) {
            rowIndex = lastRowIndex ++;
        }
        updatedRowIndicies.add(rowIndex);

        headerRow.forEach((key, colIndex) => {
            // @ts-ignore
            const value = pos[key];
            if (value) {
                sheet.getRange(rowIndex+1, colIndex+1).setValue(value);
            }
        });
    });

    const expiredCol = columnMap['expired'];
    values.forEach((row, rowIndex) => {
        if (rowIndex >= numHeaderRows) {
            sheet.getRange(rowIndex+1, expiredCol+1).setValue(updatedRowIndicies.has(rowIndex) ? null : 'EXPIRED');
        }
    });
}

function onOpen(e) {
    SpreadsheetApp.getUi()
        .createMenu('Questrade')
        .addItem('Pull', 'updatePositions')
        .addToUi();
}
