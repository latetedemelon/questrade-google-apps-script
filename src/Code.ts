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
            const ac2 = {
                accountNumber: ac.number,
                accountType: ac.type
            };

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

        const cadRates = getExchangeRates('CAD');
        const usdRates = getExchangeRates('USD');

        const symbols = [ ...getSymbols(positions.map(pos => pos.symbolId).filter(id => Number.isInteger(id))),
            { symbolId: '$CAD', currency: 'CAD' }, { symbolId: '$USD', currency: 'USD' } ];

        return positions.map(pos => {
            const symbol = symbols.find(s => s.symbolId == pos.symbolId);

            const cadExchangeRate = 1.0 / cadRates[symbol.currency];
            const valueInCAD = pos.currentMarketValue * cadExchangeRate;

            const usdExchangeRate = 1.0 / usdRates[symbol.currency];
            const valueInUSD = pos.currentMarketValue * usdExchangeRate;

            return { 
                ...pos,
                currency: symbol.currency,
                value: pos.currentMarketValue,
                valueInCAD: valueInCAD,
                valueInUSD: valueInUSD
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

    const colIndexMap = doc.getNamedRanges()
    .filter(r => r.getName().startsWith('Positions.'))
    .reduce((map, namedRange) => {
        map[namedRange.getName().replace('Positions.', '')] = namedRange.getRange().getColumn() - 1;
        return map;
    }, {});

    const accountNumCol = colIndexMap['accountNumber'];
    const symbolIdCol = colIndexMap['symbolId'];
    const updatedRowIndicies = new Set();

    positions.forEach(pos => {
        var rowIndex = values.findIndex(row => row[accountNumCol] == pos['accountNumber'] && row[symbolIdCol] == pos.symbolId);
        if (rowIndex < 0) {
            rowIndex = lastRowIndex ++;
        }
        updatedRowIndicies.add(rowIndex);

        for (const [key, colIndex] of Object.entries(colIndexMap)) {
            const value = pos[key];
            if (value) {
                //console.log(key + ' -> (' + (rowIndex+1) + ',' + (Number(colIndex)+1) + ') -> ' + value);
                sheet.getRange(rowIndex+1, Number(colIndex)+1).setValue(value);
            }
        };
    });

    const expiredCol = colIndexMap['expired'];
    values.forEach((row, rowIndex) => {
        if (rowIndex > 0) { // skip header row
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
