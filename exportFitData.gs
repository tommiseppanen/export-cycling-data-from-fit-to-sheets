function exportFitData() {
    const year = 2022;

    // Get the access token for the current user
    const accessToken = ScriptApp.getOAuthToken();

    const sessions = getCyclingSessions(accessToken, year);

    const spreadSheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = spreadSheet.getSheetByName(year.toString());

    for (let i = 0; i < sessions.session.length; i++) {
        const sessionData = getSessionData(accessToken, sessions.session[i]);
        const bucketDate = new Date(parseInt(sessions.session[i].startTimeMillis, 10));
        const durationDays = (parseInt(sessions.session[i].endTimeMillis, 10) - parseInt(sessions.session[i].startTimeMillis, 10)) / (1000 * 60 * 60 * 24);
        const avgSpeed = getValueFromDatasetPoint(sessionData.bucket[0].dataset[0], 0, 3.6);
        const maxSpeed = getValueFromDatasetPoint(sessionData.bucket[0].dataset[0], 1, 3.6);
        const minSpeed = getValueFromDatasetPoint(sessionData.bucket[0].dataset[0], 2, 3.6);
        const distance = getValueFromDatasetPoint(sessionData.bucket[0].dataset[1], 0, 1/1000);
        sheet.appendRow([bucketDate, durationDays, distance, avgSpeed, maxSpeed, minSpeed]);
    }
}

function getCyclingSessions(accessToken, year) {
    return JSON.parse(UrlFetchApp.fetch("https://www.googleapis.com/fitness/v1/users/me/sessions?startTime="+year.toString()+"-01-01T00%3A00%3A00%2B00%3A00&activityType=1&includeDeleted=true&endTime="+(year+1).toString()+"-01-01T00%3A00%3A00%2B00%3A00", {
        muteHttpExceptions: true,
        headers: {
            Authorization: "Bearer " + accessToken
        },
        "method": "get",
        "contentType": "application/json",
    }).getContentText());
}

function getSessionData(accessToken, session) {
    const request = {
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

    return JSON.parse(UrlFetchApp.fetch("https://www.googleapis.com/fitness/v1/users/me/dataset:aggregate", {
        headers: {
            Authorization: "Bearer " + accessToken
        },
        "method": "post",
        "contentType": "application/json",
        "payload": JSON.stringify(request, null, 2)
    }).getContentText());
}

function getValueFromDatasetPoint(dataset, valueIndex, multiplier) {
    if (dataset.point.length > 0) {
        return (dataset.point[0].value[valueIndex].fpVal * multiplier).toString();
    }
    return "";
}

function onOpen() {
    const ui = SpreadsheetApp.getUi();
    ui.createMenu("Google Fit")
        .addItem("Export fit data", "exportFitData")
        .addToUi();
}
