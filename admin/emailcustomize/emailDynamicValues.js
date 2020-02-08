var emailDynamicValues = [
{ "fieldVal": "{{clientLogo}}", "fieldName": "Client Logo" },
{ "fieldVal": "{{currentDateTime}}", "fieldName": "Current Date Time" },
{ "fieldVal": "{{serviceTicketNumber}}", "fieldName": "Service Ticket Number" },
{ "fieldVal": "{{accountNum}}", "fieldName": "Account Number" },
{ "fieldVal": "{{companyName}}", "fieldName": "Company Name" },
{ "fieldVal": "{{custInfo}}", "fieldName": "Customer Address" },
{ "fieldVal": "{{nameByUserNo}}", "fieldName": "Technician" },
{ "fieldVal": "{{problemDescription}}", "fieldName": "Service Notes" },
{ "fieldVal": "{{problemLocation}}", "fieldName": "Problem location" },
{ "fieldVal": "{{submittedDateTime}}", "fieldName": "Submission Date Time" },
{ "fieldVal": "{{submissionSource}}", "fieldName": "Submission Source" },
{ "fieldVal": "{{cancellationNotes}}", "fieldName": "Cancellation Notes" }
];

function displayEmailDynamicValues(listName){ 
    var ul = document.getElementById(listName);
    emailDynamicValues.sort(compare);
    for (var i = 0; i < emailDynamicValues.length; i++) {
    
      var li = document.createElement("li");
      var div1 = document.createElement("div");
      div1.appendChild(document.createTextNode(emailDynamicValues[i].fieldVal));
      var div2 = document.createElement("div");
      div2.appendChild(document.createTextNode(emailDynamicValues[i].fieldName));
      li.appendChild(div1);
      li.appendChild(div2);
      ul.appendChild(li);
    }
}

function compare(a,b) {
  if (a.fieldVal < b.fieldVal)
    return -1;
  if (a.fieldVal > b.fieldVal)
    return 1;
  return 0;
}


