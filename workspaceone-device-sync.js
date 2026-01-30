const env = 'https://xxxxxx.awmdm.com/api';
const clientId = 'XXXXXX';
const clientSecret = 'XXXXXX';
const apiKey = 'XXXXXX';
const tenantUrl ='https://emea.uemauth.vmwservices.com/connect/token';

function start(){
    writeCurrentVersions();
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Devices");
    deleteRows(sheet);
    let deviceDict = {}; //A dictionary that stores Device.Uuid and Device Serial Number

    const devicesArray = [];
    var index = 0; //Index for array
    var token = ws1_get_token();
    var responseDevice = ws1_search_devices(token);
    var devices = responseDevice.Devices;

    devices.forEach(function(device){
        //It has to be converted to valid ISO format
        let dateObj = new Date(device.LastSeen.replace(" ","T"));
        // Today's date
        let now = new Date();
        // 14 days ago
        let fourteenDaysAgo = new Date();
        fourteenDaysAgo.setDate(now.getDate() - 14);

        //Only windows devices that where turned on max 14 days ago should be written into sheet
        if(device.Platform === "WinRT" && dateObj >= fourteenDaysAgo){
            //The OS Version is splitted into two device's attributes:
            //The OS Version itself + the OS Build Version
            //They will be combined to write into column "OS Version"
            var osVersion = getOSVersion(device.OperatingSystem, device.OSBuildVersion)
            var date = device.LastSeen.split("T")[0]; //Only the date is needed
            //Now the tag of each device will be searched
            var responseTagsRaw = ws1_search_tags(token, device.Uuid)




            //Some devices have no tags, therefore the value is null/undefined
            //To prevent an exeption, the reponseTagsRaw will be parsed in JSON
            //reponseTags will either have JSON format or null
            var responseTags = safeJsonParse(responseTagsRaw);
            //Only the name of tag is needed
            //If reponseTags is null, the tagname will have the value ""
            var tagname = responseTags?.tags?.map(tag => tag.name).join(",") || "";
            //Apps per device will be searched
            var apps_response = ws1_search_apps(token, device.Uuid);
            var apps = apps_response.app_items;
            //Either they have version or are empty (if version is undefined)
            //If version is undefined, it gets the value ""
            const adobeApp = (apps.find(app => app.name === "Adobe Acrobat (64-bit)" || app.name === "Adobe Acrobat")?.installed_version) ?? "";
            const zoomApp = (apps.find(app => app.name === "Zoom" || app.name === "Zoom Workplace (64-bit)")?.installed_version) ?? "";
            const chromeApp = (apps.find(app => app.name === "GoogleChrome" || app.name === "Google Chrome")?.installed_version) ?? "";
            const forticlientApp = (apps.find(app => app.name === "FortiClient")?.installed_version) ?? "";
            const slackApp = (apps.find(app => app.name === "Slack" || app.name === "Slack (Machine)")?.installed_version) ?? "";
            const dellUpdateApp = (apps.find(app => app.name === "Dell Command | Update" || app.name === "Dell Command | Update for Windows Universal")?.installed_version) ?? "";
            const dellConfigureApp = (apps.find(app => app.name === "Dell Command | Configure")?.installed_version) ?? "";
            const dellMonitorApp = (apps.find(app => app.name === "Dell Command | Monitor")?.installed_version) ?? "";







            //the device doesn't have a tag.
            //As the Uuid is neither in sheet nor in array,
            //we use a dictionary to get Uuid by using the
            //serial number.
            deviceDict[device.SerialNumber] = device.Uuid
            //Array gets all necessary information
            devicesArray[index] =   `${device.SerialNumber};
                                    ${device.DeviceFriendlyName};
                                    ${device.UserName};
                                    ${device.Model};
                                    ${osVersion};
                                    ${date};
                                    ${device.EnrollmentStatus};
                                    ${device.ComplianceStatus};
                                    ${device.CompromisedStatus};
                                    ${tagname};
                                    ${adobeApp};
                                    ${zoomApp};
                                    ${chromeApp};
                                    ${forticlientApp};
                                    ${slackApp};
                                    ${dellUpdateApp};
                                    ${dellConfigureApp};
                                    ${dellMonitorApp}`;
                index = index + 1;
        }
    });


    //It converts an 1D array to 2D array to write into sheet
    data = devicesArray.map(item => { return item.split(";");});
    sheet.getRange(2, 1, data.length, data[0].length).setValues(data);


    //It takes the 10th column (tags)
    //This is needed to check, which rows (devices) don't have tags
    //Then tags will be written in sheet + added via POST to Workspace ONE
    var tagValues = sheet.getRange(2, 10, sheet.getLastRow() - 1, 1).getValues();
    var tagDict = getTagsID();

    for(var i = 0; i < tagValues.length; i++){
        if(tagValues[i][0] === "" || tagValues[i][0] === null){
            var deviceName = sheet.getRange(i+2, 2).getValue(); //Column 2 (DeviceFriendlyName) (e.g. AT-XXXXXX ...)
            var country = deviceName.substring(0, 2); //It gets the country's abbrevation by using the first 2 characters (e.g "AT-XXXXXX... => "AT")
            //RU belongs to PL
            if(country === "RU"){
                country = "PL";
            }

            var serialNum = sheet.getRange(i + 2, 1).getValues(); //Column 1 (SerialNumber)
            var tag_uuid = tagDict["GP" + country]; //It gets the tag UUID by using the tag's name
            var dev_uuid = deviceDict[serialNum];  //It get the Device UUID by using the Serial Number
            //If tag has a valid name:
            //It needs GP + AT/HU/PL/... which can be gotten by column 2 (eg. AT-XXXXXX)
            //However, some devices might not have the country abbrevation (e.g "DESKTOP ..." instead of "AT-DESKTOP...)
            //To avoid a script crash/exception, we only add tags which are valid via API

            if(tag_uuid){
                ws1_add_tag(token, dev_uuid, tag_uuid);
                sheet.getRange(i + 2, 10).setValue("GP" + country);
            }

        }
    }


    //Lastly, all versions which are not the newest one, will be highlighted yellow
    highlightOutdatedVersions();
}



function getZoomVersion() {
    var url = "https://support.zoom.com/hc/en/article?id=zm_kb&sysparm_article=KB0061222";
    var html = UrlFetchApp.fetch(url).getContentText(); // Find first <tr class= "table-row"> ... <td>value</td>


   var match = html.match(/<tr[^>]*table-row[^>]*>[\s\S]*?<td>(.*?)<\/td>/i);
   if (match && match[1]) {
        let result = match[1].replace(/\.\d+\s*\((\d+)\)/,'.$1');
        return result
    }

    else {
        return null;
    }
}



function getChromeVersion() {

    try {
        const url = "https://versionhistory.googleapis.com/v1/chrome/platforms/win/channels/stable/versions";
        const response = UrlFetchApp.fetch(url);
        const data = JSON.parse(response.getContentText());
        if (data.versions && data.versions.length > 0) {
            const first = data.versions[0]; // first object
            const versionJson = {
                platform: first.name, // or extract a nicer name if needed
                version: first.version
            };
            return versionJson.version;
        }
        else {
            return null;
        }
    }
        catch (e) {
            return null;
    }
}


function getFortiClientVersion(){
    var url = "https://docs.fortinet.com/product/forticlient";
    var html = UrlFetchApp.fetch(url).getContentText();
    // Regex to extract content of <title>
    var match = html.match(/<title>([\s\S]*?)<\/title>/i);
    var title = match ? match[1].trim() : null;
    title = title.split(' ')[1];
    return title;
}

function getSlackVersion(){
    var url = "https://slack.com/intl/en-gb/release-notes/windows";
    var html = UrlFetchApp.fetch(url).getContentText();
    var match = html.match(/<h2[^>]*class=["']u-flexu-align--center["'][^>]*>([\s\S]*?)<\/h2>/i);
    var version = match ? match[1].trim() : null;
    version = version.split(' ')[1]
    return version;
}


function getDellUpdateVersion(){
    const url = "https://www.dell.com/support/kbdoc/en-us/000177325/dell-command-update";
    const html = UrlFetchApp.fetch(url).getContentText();
    // Regex: find DCU X.X after "Latest Release"
    const regex = /DCU\s*([\d.]+)\s*for/i;
    const match = html.match(regex);
    if(match){
        return match[1] + ".0"; // It should have x.x.0 instead of only x.x (because 5.5 is as same as 5.5.0)
    }
    else {
        return null
    }
}

function getDellConfigureVersion(){
    const url = "https://www.dell.com/support/kbdoc/en-us/000178000/dell-command-configure";
    const html = UrlFetchApp.fetch(url).getContentText();
    // Look for "The latest version of Dell Command | Configure is" then grab X.Y.Z
    const regex =  /The latest version of Dell Command \| Configure is[^<]*<strong>v?([\d.]+)\.?<\/strong>/i;
    const match = html.match(regex);

    if (match) {
        return match[1].replace(/\.$/,""); //It removes the dot at the end (eg 5.2.0. => 5.2.0)
    }
    else {
        return null;

    }
}



function getDellMonitorVersion(){
    const url = "https://www.dell.com/support/kbdoc/en-us/000177080/dell-command-monitor";
    const html = UrlFetchApp.fetch(url).getContentText();


    // Look for "The latest version of Dell Command | Configure is" then grab X.Y.Z

    const regex = /The latest version of Dell Command \| Monitor is[^<]*<strong>v?([\d.]+)\.?<\/strong>/i;
    const match = html.match(regex);


    if (match) {
        return match[1].replace(/\.$/,"");  //It removes the dot at the end (eg 10.12.2. => 10.12.2)
    }
    else {
        return null;
    }
}



//It gets the latest versions of the apps above and writes them into sheet
function writeCurrentVersions(){
    var zoomVersion = getZoomVersion(); 
    var chromeVersion = getChromeVersion(); 
    var forticlientVersion = getFortiClientVersion();
    var slackVersion = getSlackVersion();
    var dellUpdateVersion = getDellUpdateVersion();
    var dellConfigureVersion = getDellConfigureVersion(); 
    var dellMonitorVersion = getDellMonitorVersion();       

    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Appversion");
    sheet.getRange(3, 2).setValue(zoomVersion);
    sheet.getRange(3, 3).setValue(chromeVersion);
    sheet.getRange(3, 4).setValue(forticlientVersion);
    sheet.getRange(3, 5).setValue(slackVersion);
    sheet.getRange(3, 6).setValue(dellUpdateVersion);
    sheet.getRange(3, 7).setValue(dellConfigureVersion);
    sheet.getRange(3, 8).setValue(dellMonitorVersion);
}


/*
THESE FUNCTIONS SERVE AS A HELP FOR THE MAIN FUNCTION:
INSTEAD OF WRITING THE WHOLE CODE IN MAIN FUNCTION,
THE CERTAIN FUNCTION WILL BE CALLED.
IT MAKES THE MAIN FUNCTION CLEARER.
*/
//It deletes all rows (except the first one) before the devices will bewritten into sheet
function deleteRows(sheet){
    var lastRow = sheet.getLastRow();
    var lastColumn = sheet.getLastColumn();
    if (lastRow > 1) { // only clear if there are more than 1 rows
        var range = sheet.getRange(2, 1, lastRow - 1, lastColumn);
        range.clearContent(); // clears all rows except row 1
    }
}



//It highlights all outdated versions yellow
function highlightOutdatedVersions(){
    var devices_sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Devices");
    var appversion_sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Appversion");
    var adobeVersion = appversion_sheet.getRange(3, 1).getValue();
    var zoomVersion = appversion_sheet.getRange(3, 2).getValue();
    var chromeVersion = appversion_sheet.getRange(3, 3).getValue();

    var forticlientVersion = appversion_sheet.getRange(3, 4).getValue();
    var slackVersion = appversion_sheet.getRange(3, 5).getValue();
    var dellUpdateVersion = appversion_sheet.getRange(3, 6).getValue();
    var dellConfigureVersion = appversion_sheet.getRange(3, 7).getValue();
    var dellMonitorVersion = appversion_sheet.getRange(3, 8).getValue();
    var row = 2;

    while(row <= devices_sheet.getLastRow()){
        var adobeTemp = devices_sheet.getRange(row, 11).getValue();
        var zoomTemp = devices_sheet.getRange(row, 12).getValue();
        var chromeTemp = devices_sheet.getRange(row, 13).getValue();
        var forticlientTemp = devices_sheet.getRange(row, 14).getValue();
        var slackTemp = devices_sheet.getRange(row, 15).getValue();
        var dellUpdateTemp = devices_sheet.getRange(row, 16).getValue();
        var dellConfigureTemp = devices_sheet.getRange(row, 17).getValue();
        var dellMonitorTemp = devices_sheet.getRange(row, 18).getValue();
        
        if(adobeTemp !== adobeVersion){
            devices_sheet.getRange(row, 11).setBackground("yellow");
        }
    
        if(zoomTemp !== zoomVersion){
            devices_sheet.getRange(row, 12).setBackground("yellow");
        }

        if(chromeTemp !== chromeVersion){
            devices_sheet.getRange(row, 13).setBackground("yellow");
        }

        if(forticlientTemp !== forticlientVersion){
            devices_sheet.getRange(row, 14).setBackground("yellow");
        }


        if(slackTemp !== slackVersion){
            devices_sheet.getRange(row, 15).setBackground("yellow");
        }

        if(dellUpdateTemp !== dellUpdateVersion){
            devices_sheet.getRange(row, 16).setBackground("yellow");
        }

        if(dellConfigureTemp !== dellConfigureVersion){
            devices_sheet.getRange(row, 17).setBackground("yellow");
        }

        if(dellMonitorTemp.slice(0,-2) !== dellMonitorVersion){
            devices_sheet.getRange(row, 18).setBackground("yellow");
        }

        row = row + 1;
    }
}

//It returns a dictionary of all 9 Tags and their Uuid
function getTagsID(){
    tagIDs = {
        "GPAT" : "XXXXXX",
        "GPHR" : "XXXXXX",
        "GPHU" : "XXXXXX",
        "GPPL" : "XXXXXX",
        "GPBG" : "XXXXXX",
        "GPUA" : "XXXXXX",
        "GPSK" : "XXXXXX",
        "GPSI" : "XXXXXX",
        "GPRO" : "XXXXXX",

    }
    return tagIDs;
}


//Parses text to JSON format
//If text is null/undefined, it returns null
function safeJsonParse(text){
    try{
        return JSON.parse(text)
    }
    catch(e){
        return null;
    }
}


//The function puts Operating System (Version Number) and Build Versiontogether (For example: "10.0.26100" + "4946" => "26100.4946")
function getOSVersion(operatingSystem, buildVersion)
{
    operatingSystem = String(operatingSystem).substring(5);
    osVersion = operatingSystem + "." + buildVersion;
    return osVersion;
}



/*
THESE FUNCTIONS GET TOKEN, TAGS, DEVICES
AND APPS OR ADD TAG TO DEVICE VIA API.
THEY ARE USED BY THE MAIN FUNCTION.
*/
function ws1_get_token() {
    const authHeader = Utilities.base64Encode(`${clientId}:${clientSecret}`);
    const tokenResponse = UrlFetchApp.fetch(`${tenantUrl}`, {
            method: 'post',
            headers: {
                'Authorization': `Basic ${authHeader}`,
                'Content-Type': 'application/x-www-form-urlencoded'
            },
        payload: {
            'grant_type': 'client_credentials'
        }
    });
    return token = JSON.parse(tokenResponse.getContentText()).access_token;
}
//It gets tags for the specific device
function ws1_search_tags(token, dev_uuid){
    var response = UrlFetchApp.fetch(`${env}/mdm/devices/${dev_uuid}/tags`,{
        method: 'get',
        headers: {
            'Accept': 'application/json',
            'Authorization': `Bearer ${token}`,
            'aw-tenant-code': `${apiKey}`
        }
    });
    var tags_rq = response.getContentText();
    return tags_rq;
}


//It gets all devices
function ws1_search_devices(token){
    var devices_rq = JSON.parse(UrlFetchApp.fetch(`${env}/mdm/devices/search`,{
        method: 'get',
        headers: {
            'Accept': 'application/json',
            'Authorization': `Bearer ${token}`,
            'aw-tenant-code': `${apiKey}`
        }
    }));
    return devices_rq;
}


//It gets apps for the specific device
function ws1_search_apps(token,dev_uuid){
    var software_rq = JSON.parse(UrlFetchApp.fetch(`${env}/mdm/devices/${dev_uuid}/apps/search`, {
        method: 'get',
        headers: {
            'Accept': 'application/json',
            'Authorization': `Bearer ${token}`,
            'aw-tenant-code': `${apiKey}`

    }}));
    return software_rq;
}

