function reset() {
    const service = getService();
    service.reset();
}

function getService() {
    // @ts-ignore
    return OAuth2.createService("questrade")
        .setAuthorizationBaseUrl("https://login.questrade.com/oauth2/authorize")
        .setTokenUrl("https://login.questrade.com/oauth2/token")
        .setClientId(PropertiesService.getScriptProperties().getProperty("consumerKey"))
        .setClientSecret("secret") // No secret provided by QT. Use dummy one to make oauth2 lib happy.
        .setCallbackFunction("authCallback")
        .setPropertyStore(PropertiesService.getUserProperties())
        .setScope("read_acc")
        .setParam("response_type", "code");
}

function authCallback(request) {
    const service = getService();
    const authorized = service.handleCallback(request);
    if (authorized) {
        return HtmlService.createHtmlOutput(
            "Success! <script>setTimeout(function() { top.window.close() }, 1);</script>"
        );
    } else {
        return HtmlService.createHtmlOutput("Denied.");
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
            "Pull again when the authorization is complete."
        );
        // @ts-ignore
        template.authorizationUrl = authorizationUrl;
        const page = template.evaluate();
        SpreadsheetApp.getUi().showSidebar(page);
        return;
    }

    const apiOptions: GoogleAppsScript.URL_Fetch.URLFetchRequestOptions = {
        method: "get",
        headers: {
            Authorization: "Bearer " + service.getAccessToken(),
        },
    };
    const apiUrl = service.getToken().api_server;

    const fetch = function (query) {
        return JSON.parse(UrlFetchApp.fetch(apiUrl + query, apiOptions).getContentText());
    };

    const getAccounts = (this.getAccounts = function () {
        return fetch("v1/accounts").accounts;
    });
    this.accounts = this.getAccounts();

    const getPositions = (this.getPositions = function (accountNumber) {
        return fetch("v1/accounts/" + accountNumber + "/positions").positions;
    });

    const getSymbols = (this.getSymbols = function (symbolIds) {
        return fetch("v1/symbols?ids=" + symbolIds.join(",")).symbols;
    });

    const getBalances = (this.getBalances = function (accountNumber) {
        return fetch("v1/accounts/" + accountNumber + "/balances").perCurrencyBalances;
    });

    const getOrders = (this.getOrders = function (accountNumber) {
        return fetch("v1/accounts/" + accountNumber + "/orders"
            + "?startTime=1999-12-31T00:00:00.000000Z"
            + "&stateFilter=Open&").orders;
    });

    const getCandles = (this.getCandles = function (symbolId, startTime) {
        return fetch("v1/markets/candles/" + symbolId
            + "?startTime=" + startTime
            + "&endTime=2099-12-31T00:00:00.000000Z"
            + "&interval=OneYear")
            .candles;
    });

    const getHighSince = (this.getHighSince = function (symbolId, startTime) {
        try {
            var highs = getCandles(symbolId, startTime).map(c => c.high);
            return Math.max(...highs);
        } catch (err) {
            console.log("Error pulling candles for symbolId " + symbolId, err);
            return undefined;
        }
    });

    const getExchangeRates = (base) => {
        return JSON.parse(UrlFetchApp.fetch(`https://api.exchangeratesapi.io/latest?base=${base}`).getContentText())
            .rates;
    };

    const prefixKeys = (map, prefix) => {
        if (map) {
            return Object.entries(map).reduce((m, [k, v]) => {
                m[prefix + k] = v;
                return m;
            }, {});
        }
    };

    const getDuration = (order) => {
        switch (order.timeInForce) {
            case "GoodTillCanceled":
                return "GTC";
            default:
                return "?";
        }
    };

    const getEnrichedPositions = (this.getEnrichedPositions = function () {
        const positions = getAccounts().flatMap((ac) => {
            const ac2 = {
                accountNumber: ac.number,
                accountType: ac.type,
            };

            const acPositions = getPositions(ac.number).map((pos) => ({
                ...pos,
                ...ac2,
            }));

            getBalances(ac.number).forEach((bal) => {
                acPositions.push({
                    symbol: "$" + bal.currency,
                    symbolId: "$" + bal.currency,
                    currentMarketValue: bal.cash,
                    ...ac2,
                });
            });

            const orderMap = getOrders(ac.number)
                .filter((order) => order.state == "Accepted" || order.state == "Queued")
                .reduce((map, order) => {
                    var orders = map[order.symbolId];
                    if (!orders) {
                        orders = [];
                        map[order.symbolId] = orders;
                    }
                    orders.push(order);
                    return map;
                }, {});
            // console.log('orderMap: ', orderMap);

            acPositions.forEach((pos) => {
                const orders = orderMap[pos.symbolId] || [];

                pos.sharesNotStopped = orders.length > 0 ? "?" : null;
                pos.stop = null;
                pos.limit = null;
                pos.duration = null;
                pos.high = null;

                if (orders.length == 1) {
                    const order = orders[0];
                    if (order.side == "Sell") {
                        switch (order.orderType) {
                            case "TrailStopLimitInPercentage":
                                pos.sharesNotStopped = pos.openQuantity - order.totalQuantity;
                                pos.stop = order.stopPrice + "%";
                                pos.limit = order.limitPrice + (Boolean(order.isLimitOffsetInDollar) ? "" : "%");
                                pos.duration = getDuration(order);
                                pos.high = getHighSince(order.symbolId, order.updateTime);
                                break;
                            case "TrailStopInPercentage":
                                pos.sharesNotStopped = pos.openQuantity - order.totalQuantity;
                                pos.stop = order.stopPrice + "%";
                                pos.duration = getDuration(order);
                                pos.high = getHighSince(order.symbolId, order.updateTime);
                                break;
                        }
                    }
                }
            });

            return acPositions;
        });

        const cadRates = getExchangeRates("CAD");
        const usdRates = getExchangeRates("USD");

        const symbols = [
            ...getSymbols(positions.map((pos) => pos.symbolId).filter((id) => Number.isInteger(id))),
            { symbolId: "$CAD", currency: "CAD" },
            { symbolId: "$USD", currency: "USD" },
        ];

        return positions.map((pos) => {
            const symbol = symbols.find((s) => s.symbolId == pos.symbolId);

            const cadExchangeRate = 1.0 / cadRates[symbol.currency];
            const valueInCAD = pos.currentMarketValue * cadExchangeRate;

            const usdExchangeRate = 1.0 / usdRates[symbol.currency];
            const valueInUSD = pos.currentMarketValue * usdExchangeRate;

            return {
                ...pos,
                currency: symbol.currency,
                value: pos.currentMarketValue,
                valueInCAD: valueInCAD,
                valueInUSD: valueInUSD,
            };
        });
    });
};

function updatePositions() {
    const qt = new QuestradeApiSession();
    const positions = qt.getEnrichedPositions();

    const doc = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = doc.getSheetByName("Positions");
    const values = sheet.getDataRange().getValues();
    var lastRowIndex = values.length;

    const colIndexMap = doc
        .getNamedRanges()
        .filter((r) => r.getName().startsWith("Positions."))
        .reduce((map, namedRange) => {
            map[namedRange.getName().replace("Positions.", "")] = namedRange.getRange().getColumn() - 1;
            return map;
        }, {});

    const accountNumCol = colIndexMap["accountNumber"];
    const symbolIdCol = colIndexMap["symbolId"];
    const symbolCol = colIndexMap["symbol"];
    const updatedRowIndicies = new Set();

    positions
        .filter((pos) => typeof pos.currentMarketValue == "number")
        .forEach((pos) => {
            var rowIndex = values.findIndex(
                (row) => row[accountNumCol] == pos["accountNumber"] &&
                    (row[symbolIdCol] == pos.symbolId || (!row[symbolIdCol] && row[symbolCol] == pos.symbol))
            );
            if (rowIndex < 0) {
                rowIndex = lastRowIndex++;
            }
            updatedRowIndicies.add(rowIndex);

            for (const [key, colIndex] of Object.entries(colIndexMap)) {
                const value = pos[key];
                if (value !== undefined) {
                    sheet.getRange(rowIndex + 1, Number(colIndex) + 1).setValue(value);
                }
            }
        });

    const statusCol = colIndexMap["status"];
    const expiredNullCols = [
        "value",
        "currentMarketPrice",
        "valueInCAD",
        "valueInUSD",
        "stop",
        "limit",
        "sharesNotStopped",
        "duration",
        "high"
    ]
        .map((key) => colIndexMap[key])
        .filter((col) => col !== undefined);

    values.forEach((row, rowIndex) => {
        if (rowIndex > 0) {
            if (updatedRowIndicies.has(rowIndex)) {
                sheet.getRange(rowIndex + 1, statusCol + 1).setValue(null);
            } else {
                sheet.getRange(rowIndex + 1, statusCol + 1).setValue("🛑");
                expiredNullCols.forEach((colIndex) => sheet.getRange(rowIndex + 1, colIndex + 1).setValue(null));
            }
        }
    });
}

function sortByNamedRange(name, ascending) {
    const doc = SpreadsheetApp.getActiveSpreadsheet();
    const range = doc.getRangeByName(name);
    range.getSheet().sort(range.getColumn(), ascending);
}

function sortBySortId() {
    sortByNamedRange("Positions.sortId", true);
};

function sortByRebalanceAmount() {
    sortByNamedRange("Positions.rebalanceAmount", false);
};

function onOpen(e) {
    SpreadsheetApp.getUi()
        .createMenu("STONKS!")
        .addItem("Pull from Questrade", "updatePositions")
        .addItem("Sort by Sort Id", "sortBySortId")
        .addItem("Sort by Rebal $", "sortByRebalanceAmount")
        .addToUi();
}
