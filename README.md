**`fetch-outlook-mail`** is a module to fetch mail from Professional or Academic Outlook email accounts.

Uses Microsoft Outlook Mail REST API. First-time users must allow one-time app access at: https://login.microsoftonline.com/common/oauth2/v2.0/authorize?response_type=code&client_id=0a64808c-eeb9-40bf-8554-4e9d22791831&redirect_uri=http%3A%2F%2Flocalhost%3A3000%2Fauthorize&scope=openid%20profile%20offline_access%20User.Read%20Mail.Read&sso_reload=true

Copyright (c) Microsoft. All rights reserved.

## Installation

```javascript
npm install fetch-outlook-mail
```

## Example

```javascript
var Fetch = require('fetch-outlook-mail')

let fetcher = new Fetch('username', 'password');

function logInfo(data) {
  // The following functions return values describing the mailbox as a whole
  let total = fetcher.totalQuantity(data);
  let totalMail = fetcher.allMail(data);
  let allFolders = fetcher.folderList(data);

  // The following get more specific; filtering through the data
  let mailFromNathan = fetcher.fromQuantity(data, "sample.address@foo.com", "Inbox");
  let mailToNathan = fetcher.toQuantity(data, "sample.address@foo.com", "Sent Items");
  let yesterdaysEmail = fetcher.dateTimeMail(data, "2018-07-09", "00:00:00", "2018-07-14", "00:00:00", "Inbox");
  let errorEmails = fetcher.bodyContainsString(data, "invalid uid", "Inbox");
}

fetcher.logInCreateData(logInfo);
```

## Exhaustive Documentation 
```javascript
/** 
     * Creates fetcher object, generates access token options
     * @param username login credential email id
     * @param password login credential password
     * @return none
     */
    constructor(username, password)

    get username()
    get password()
    get tokenOptions() 

    set username(value) 
    set password(value) 

    /**
     * Retrieves access token and fetches data from Outlook API 
     * @param callback callback function for desired data manipulation
     * @return none
     */
    logInCreateData(callback)

    /**
     * @param data dictionary object created by logInCreateData
     * @return array of folders in Outlook account
     */
    folderList(data)

    /**
     * @param data dictionary object created by logInCreateData
     * @return dictionary with name and email address associated with account
     */
    profileInfo(data) 

    /**
     * @param data dictionary object created by logInCreateData
     * @returns int total amount of emails in account OR 1000 if API is maxed out
     */
    totalQuantity(data) 

    /**
     * @param data dictionary object created by logInCreateData
     * @return array of all mail in every folder
     */
    allMail(data) 

    /**
     * @param data dictionary object created by logInCreateData
     * @param desiredFolder string containing folder to search in 
     * @return boolean if desired folder exists in Outlook folders list
     */
    folderExists(data, desiredFolder) 

    /**
     * @param data dictionary object created by logInCreateData
     * @param desiredFolder string containing folder to search in
     * @return int total count of folders associated with account
     */
    folderQuantity(data, desiredFolder)

    /**
     * @param data dictionary object created by logInCreateData
     * @param desiredFolder string containing folder to search in
     * @return array of all mail that exists in specified folder
     */
    folderMail(data, desiredFolder)

    /**
     * @param data dictionary object created by logInCreateData 
     * @param to string specifying email address in to field
     * @param desiredFolder string containing folder to search in
     * @return int of all mail to specific email in desired search folder
     */
    toQuantity(data, to, desiredFolder) 

    /**
     * @param data dictionary object created by logInCreateData 
     * @param to string specifying email address in to field
     * @param desiredFolder string containing folder to search in
     * @return array of all mail objects to specific email in desired search folder
     */
    toMail(data, to, desiredFolder) 

    /**
     * @param data dictionary object created by logInCreateData 
     * @param from string specifying email address in from field
     * @param desiredFolder string containing folder to search in
     * @return int of all mail from specific email in desired search folder
     */
    fromQuantity(data, from, desiredFolder) 

    /**
     * @param data dictionary object created by logInCreateData 
     * @param from string specifying email address in from field
     * @param desiredFolder string containing folder to search in
     * @return array of all mail from specific email in desired search folder
     */
    fromMail(data, from, desiredFolder) 

    /**
     * @param data dictionary object created by logInCreateData 
     * @param subject string specifying subject match; is not case or empty space sensitive
     * @param desiredFolder string containing folder to search in
     * @return int of all mail with specific subject name in desired search folder
     */
    subjectQuantity(data, subject, desiredFolder) 

    /**
     * @param data dictionary object created by logInCreateData 
     * @param subject string specifying subject match; is not case or empty space sensitive
     * @param desiredFolder string containing folder to search in
     * @return array of all mail with specific subject name in desired search folder
     */
    subjectMail(data, subject, desiredFolder) 

    /**
     * @param data dictionary object created by logInCreateData 
     * @param startDate string "YYYY-MM-DD"
     * @param startTime string "HH:MM:SS"
     * @param finishDate string "YYYY-MM-DD"
     * @param finishTime string "HH:MM:SS"
     * @param desiredFolder string containing folder to search in
     * @return int total of emails received between desired time frames in desired search folder
     */
    dateTimeQuantity(data, startDate, startTime, finishDate, finishTime, desiredFolder) 

    /**
     * @param data dictionary object created by logInCreateData 
     * @param startDate string "YYYY-MM-DD"
     * @param startTime string "HH:MM:SS"
     * @param finishDate string "YYYY-MM-DD"
     * @param finishTime string "HH:MM:SS"
     * @param desiredFolder string containing folder to search in
     * @return array of emails received between desired time frames in desired search folder
     */
    dateTimeMail(data, startDate, startTime, finishDate, finishTime, desiredFolder) 

    /**
     * @param data dictionary object created by logInCreateData 
     * @param str string containing search string
     * @param desiredFolder string containing folder to search in
     * @return array containing all emails with bodies that have the desired search string in them
     */
    bodyContainsString(data, str, desiredFolder) 
```