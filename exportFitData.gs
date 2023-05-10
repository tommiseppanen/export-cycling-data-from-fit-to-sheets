function exportFitData() {
    const year = 2022;

    // Get the access token for the current user
    let accessToken = ScriptApp.getOAuthToken();

    let sessions = getSessions(accessToken, year);

    let spreadSheet = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = spreadSheet.getSheetByName(year.toString());

    for (let i = 0; i < sessions.session.length; i++) {
        let sessionData = getSessionData(accessToken, sessions.session[i]);

        let bucketDate = new Date(parseInt(sessions.session[i].startTimeMillis, 10));
        let durationDays = (parseInt(sessions.session[i].endTimeMillis, 10) - parseInt(sessions.session[i].startTimeMillis, 10)) / (1000 * 60 * 60 * 24);
        let avgSpeed = -1;
        let maxSpeed = -1;
        let minSpeed = -1;
        let distance = -1;

        if (sessionData.bucket[0].dataset[0].point.length > 0) {
            avgSpeed = sessionData.bucket[0].dataset[0].point[0].value[0].fpVal * 3.6;
        }

        if (sessionData.bucket[0].dataset[0].point.length > 0) {
            maxSpeed = sessionData.bucket[0].dataset[0].point[0].value[1].fpVal * 3.6;
        }

        if (sessionData.bucket[0].dataset[0].point.length > 0) {
            minSpeed = sessionData.bucket[0].dataset[0].point[0].value[2].fpVal * 3.6;
        }

        if (sessionData.bucket[0].dataset[1].point.length > 0) {
            distance = sessionData.bucket[0].dataset[1].point[0].value[0].fpVal / 1000;
        }

        sheet.appendRow([bucketDate, durationDays,
            distance == -1 ? ' ' : distance,
            avgSpeed == -1 ? ' ' : avgSpeed,
            maxSpeed == -1 ? ' ' : maxSpeed,
            minSpeed == -1 ? ' ' : minSpeed]);
    }
}

function getSessions(accessToken, year) {
    return JSON.parse(UrlFetchApp.fetch('https://www.googleapis.com/fitness/v1/users/me/sessions?startTime='+year.toString()+'-01-01T00%3A00%3A00%2B00%3A00&activityType=1&includeDeleted=true&endTime='+(year+1).toString()+'-01-01T00%3A00%3A00%2B00%3A00', {
        muteHttpExceptions: true,
        headers: {
            Authorization: 'Bearer ' + accessToken
        },
        'method': 'get',
        'contentType': 'application/json',
    }).getContentText());
}

function getSessionData(accessToken, session) {
    let request = {
        "aggregateBy": [
            {
                "dataTypeName": "com.google.speed"
            },
            {
                "dataTypeName": "com.google.distance.delta"
            }
        ],
        "bucketBySession": {},
        "startTimeMillis": session.startTimeMillis,
        "endTimeMillis": session.endTimeMillis
    };

    return JSON.parse(UrlFetchApp.fetch('https://www.googleapis.com/fitness/v1/users/me/dataset:aggregate', {
        headers: {
            Authorization: 'Bearer ' + accessToken
        },
        'method': 'post',
        'contentType': 'application/json',
        'payload': JSON.stringify(request, null, 2)
    }).getContentText());
}

function onOpen() {
    let ui = SpreadsheetApp.getUi();
    ui.createMenu('Google Fit')
        .addItem('Export fit data', 'exportFitData')
        .addToUi();
}
