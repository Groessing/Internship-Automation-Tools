function main(){
    var candidatesList = readFromCSV();
    candidatesList = addCandidateToRecruitee(candidatesList);
    candidatesList = setFieldsToRecruitee(candidatesList);
    setDisqualify(candidatesList);
}

function readFromCSV()
{
    var candidatesList = [];
    var sheet = SpreadsheetApp.getActiveSheet();
    var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();

    for(let i = 0; i < data.length; i++){
        var row = data[i];
        var birthdayString = row[5];
        var regristrationString = row[10];
        var name = row[0] + " " + row[1];
        var birthday = parseDate(birthdayString)
        var regristration = parseDate(regristrationString)
        candidatesList[i] = `${name};
                            ${row[2]};
                            ${row[3]};
                            ${row[4]};
                            ${birthday};
                            ${regristration};
                            ${row[17]};
                            ${row[46]};
                            ${row[47]}`;
    }
    return candidatesList;
}


//Parses DD.MM.YYYY to YYYY-MM-DD
function parseDate(dateString)
{
    const parts = dateString.split(".");    
    return `${parts[2]}-${parts[1]}-${parts[0]}`;
}


//Add candidates to Recruitee and stores their IDs to the candidatesList
function addCandidateToRecruitee(candidatesList)
{
    const apiToken = 'XXXXXX';
    const companyId = '12345';

    for(let i = 0; i < candidatesList.length; i++)
    {
        var split = candidatesList[i].split(";");
        const candidate = {
            offer_id: 1234567,
            candidate:
            {
                name: split[0],
                emails: [split[2]],
                phones: [split[1]],
                tags: ["onlyfy"]
            }
        };

        const url = `https://api.recruitee.com/c/${companyId}/candidates`;

        const options = {
            method: "post",
            contentType: "application/json",
            headers: {
                Authorization: `Bearer ${apiToken}`
            },
            payload: JSON.stringify(candidate),
            muteHttpExceptions: true
        };
        const response = UrlFetchApp.fetch(url, options);
        
        const status = response.getResponseCode();
        const body = response.getContentText();
        const json = JSON.parse(body);
        
        const candidateId = json?.candidate?.id;
        candidatesList[i] = `${candidatesList[i]};${candidateId}`;
    }
    return candidatesList;
}


//Sets fields to candidates after they were added (address, birthday, regristration, UTM Source)
function setFieldsToRecruitee(candidatesList)
{
    const apiToken = 'XXXXXX';
    const companyId = '12345';


    for(let i = 0; i < candidatesList.length; i++)
    {
        var split = candidatesList[i].split(';');
        candidateId = split[9];
        const url = `https://api.recruitee.com/c/${companyId}/custom_fields/candidates/${candidateId}/fields`;

        const fields = [
        {
            kind: 'date_of_birth',
            values: [{ date: split[4]}]
        },
        {
            name: 'UTM Source',
            kind: 'single_line',
            values: [{text: split[6]}]
        },
        {
            name: 'Address',
            kind: 'single_line',
            values: [{text: split[3]}]
        },
        {
            name: 'Registration Date',
            kind: 'date',
            values: [{date: split[5]}]
        }
        ];

        for(const fieldId in fields)
        {
            const options = {
                method: 'POST',
                contentType: 'application/json',
                headers: {
                    Authorization: `Bearer ${apiToken}`
                },
                payload: JSON.stringify({field: fields[fieldId]}),
                muteHttpExceptions: true
                };
                const response = UrlFetchApp.fetch(url, options);
        }
    }
    return candidatesList;
}

//Sets Disqualify to candidates according to the disqualify ID
function setDisqualify(candidatesList)
{
    const apiToken = 'XXXXXX';
    const companyId = '12345';
    const url = `https://api.recruitee.com/c/${companyId}/bulk/candidates/disqualify`;

    for(let i = 0; i < candidatesList.length; i++){
        var split = candidatesList[i].split(';');
        var candidateId = split[9];
        var disqualifyReasonId = split[8];
        
        const payload = {
            disqualify_reason_id: disqualifyReasonId,
            candidates: [candidateId],
            offer_id: 1234567
        };

        const options = {
            method: 'PATCH',
            contentType: 'application/json',
            headers: {
                Authorization: `Bearer ${apiToken}`
            },
            payload: JSON.stringify(payload),
            muteHttpExceptions: true
        };
        const response = UrlFetchApp.fetch(url, options);
    }
}