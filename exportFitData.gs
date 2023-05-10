function exportFitData() {

  // Get the access token for the current user
  var accessToken = ScriptApp.getOAuthToken();


  var response = UrlFetchApp.fetch('https://www.googleapis.com/fitness/v1/users/me/sessions?startTime=2022-01-01T00%3A00%3A00%2B00%3A00&activityType=1&includeDeleted=true&endTime=2023-01-01T00%3A00%3A00%2B00%3A00', {
    muteHttpExceptions: true,
    headers: {
      Authorization: 'Bearer ' + accessToken
    },
    'method' : 'get',
    'contentType' : 'application/json',
  });

  //Logger.log(response.getContentText()); 

  var json = JSON.parse(response.getContentText());

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("2022");

  for(var b = 0; b < json.session.length; b++) {


  var request = {
      "aggregateBy": [
        {
          "dataTypeName": "com.google.speed"
        },
        {
          "dataTypeName": "com.google.distance.delta"
        }
      ],
      "bucketBySession": {},
      "startTimeMillis": json.session[b].startTimeMillis,
      "endTimeMillis": json.session[b].endTimeMillis
    };

var response2 = UrlFetchApp.fetch('https://www.googleapis.com/fitness/v1/users/me/dataset:aggregate', {
    headers: {
      Authorization: 'Bearer ' + accessToken
    },
    'method' : 'post',
    'contentType' : 'application/json',
    'payload' : JSON.stringify(request, null, 2)
  });
  
   var json2 = JSON.parse(response2.getContentText());

    var bucketDate = new Date(parseInt(json.session[b].startTimeMillis, 10));
    var durationDays = (parseInt(json.session[b].endTimeMillis, 10)-parseInt(json.session[b].startTimeMillis, 10))/(1000*60*60*24);
    var avgSpeed = -1;
    var maxSpeed = -1;
    var minSpeed = -1;
    var distance = -1;
    
    if (json2.bucket[0].dataset[0].point.length > 0) {
      avgSpeed = json2.bucket[0].dataset[0].point[0].value[0].fpVal*3.6;
    }
    
    if (json2.bucket[0].dataset[0].point.length > 0) {
      maxSpeed = json2.bucket[0].dataset[0].point[0].value[1].fpVal*3.6;
    }

    if (json2.bucket[0].dataset[0].point.length > 0) {
      minSpeed = json2.bucket[0].dataset[0].point[0].value[2].fpVal*3.6;
    }

    if (json2.bucket[0].dataset[1].point.length > 0) {
      distance = json2.bucket[0].dataset[1].point[0].value[0].fpVal/1000;
    }

    
    sheet.appendRow([bucketDate, durationDays,
                    distance == -1 ? ' ' : distance, 
                     avgSpeed == -1 ? ' ' : avgSpeed, 
                     maxSpeed == -1 ? ' ' : maxSpeed, 
                     minSpeed == -1 ? ' ' : minSpeed]);
    }
}

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Google Fit')
      .addItem('Export fit data', 'exportFitData')
      .addToUi();
}
