const request = require('request');
const htmlToText = require('html-to-text');

// These are constant, since they are required and unchanging properties of the client itself.
const CLIENT_ID = 'PLEASE PROVIDE -- SEE README FOR CLIENT SETUP INSTRUCTIONS';
const CLIENT_SECRET = 'PLEASE PROVIDE -- SEE README FOR CLIENT SETUP INSTRUCTIONS ';
const MAX = 1000;

class Fetch {
    /** 
     * Creates fetcher object, generates access token options
     * @param username login credential email id
     * @param password login credential password
     * @return none
     */
    constructor(username, password) {
        if (!username || !password) {
            throw ('Invalid object initialization - must provide username and password');
        }
        this._username = username; //set immediately, since required for data fetching
        this._password = password;
        this._tokenOptions = null;
        this._generateTokenOptions();
    }

    get username() { return this._username; }

    get password() { return this._password; }

    get tokenOptions() { return this._tokenOptions; }

    set username(value) {
        this._username = value;
    }

    set password(value) {
        this._password = value;
    }

    /** 
     * Helper function for constructor that generates token options based on login credentials
     * @param none
     * @return none
     */
    _generateTokenOptions() { 
        this._tokenOptions = {
            url: 'https://login.windows.net/common/oauth2/token', 
            headers: { 'content-type': 'application/x-www-form-urlencoded' },
            body: `resource=https%3A%2F%2Fgraph.microsoft.com%2F&grant_type=password&username=${this.username}&password=${this.password}&audience=https://someapi.com/api&scope=openid&client_id=${CLIENT_ID}&client_secret=${CLIENT_SECRET}`,
            json: true };
    }

    /**
     * Retrieves access token and fetches data from Outlook API 
     * @param callback callback function for desired data manipulation
     * @return none
     */
    logInCreateData(callback) {  
        let data = {};
        let messages = [];
        request(this.tokenOptions, function (error, response, tokenObj) {
            if (tokenObj.error) {
                console.log(tokenObj);
                throw ('Error - Invalid login credentials');
            }
            var fetchOptions = { method: 'GET',
                url: `https://outlook.office.com/api/v2.0/me/messages?$top=${MAX}`,
                headers:
                { authorization: `Bearer ${tokenObj.access_token}`,
                'content-type': 'application/json' },
                json: true
            };
            request(fetchOptions, function (error, response, mailObj) {
                if (error) throw new Error(error);
                for (let i = (mailObj.value.length-1); i >= 0; i--) {
                    let message = {};
                    message['emailNumber'] = i+1;
                    message['folderId'] = mailObj.value[i].ParentFolderId;
                    message['timestamp'] = mailObj.value[i].ReceivedDateTime;
                    // handle to field
                    var to = [];
                    for (let x = 0; x < mailObj.value[i].ToRecipients.length; x++) {
                        to.push(mailObj.value[i].ToRecipients[x].EmailAddress.Address);
                    }
                    message['to'] = to;
                    // handle from and subject fields
                    if (mailObj.value[i].From !== undefined) {
                        message['from'] = mailObj.value[i].From.EmailAddress.Address;   
                    } else {
                        message['from'] = 'draft';
                    }
                    var subject = mailObj.value[i].Subject;
                    if (subject === '') {subject = 'None';}
                    message['subject'] = subject;
                    // handle body field
                    var bodyhtml = mailObj.value[i].Body.Content;
                    var bodytext = htmlToText.fromString(bodyhtml, {
                        wordwrap: 130
                    });
                    bodytext = bodytext.split('From')[0];
                    if (bodytext === '') {bodytext = 'None';}
                    message['body'] = bodytext;
                    messages.push(message);
                }
                data.messages = messages;
                fetchOptions['url'] = 'https://outlook.office.com/api/v2.0/me/MailFolders';
                request(fetchOptions, function (error, response, folderObj) {
                    if (error) throw new Error(error);
                    let folders = [];
                    for (let i = 0; i < folderObj.value.length; i++) {
                        let folder = {};
                        folder.name = folderObj.value[i].DisplayName;
                        folder.id = folderObj.value[i].Id;
                        folders.push(folder);
                    }
                    data.folders = folders;
                    fetchOptions['url'] = 'https://outlook.office.com/api/v2.0/me';
                    request(fetchOptions, function (error, response, profileObj) {
                        if (error) throw new Error(error);
                        let profile = {};
                        profile['Display Name'] = profileObj.DisplayName;
                        profile['Email Address'] = profileObj.EmailAddress;
                        data.profile = profile;
                        callback(data);
                    });
                 });
            });   
        });
    }

    /**
     * @param data dictionary object created by logInCreateData
     * @return array of folders in Outlook account
     */
    folderList(data) {  
        return data.folders;
    }

    /**
     * @param data dictionary object created by logInCreateData
     * @return dictionary with name and email address associated with account
     */
    profileInfo(data) { 
        return data.profile;
    }

    /**
     * @param data dictionary object created by logInCreateData
     * @returns int total amount of emails in account OR 1000 if API is maxed out
     */
    totalQuantity(data) {  
        if (data.messages.length < MAX) {
            return data.messages.length;
        }
        else {
            console.log(`WARNING: totalQuantity is maxed out at ${MAX} messages`);
            return MAX;
        }
    }

    /**
     * @param data dictionary object created by logInCreateData
     * @return array of all mail in every folder
     */
    allMail(data) {  
        return data.messages;
    }

    /**
     * @param data dictionary object created by logInCreateData
     * @param desiredFolder string containing folder to search in 
     * @return boolean if desired folder exists in Outlook folders list
     */
    folderExists(data, desiredFolder) {  
        let folderIndex = -1;
        for(let i = 0; i < data.folders.length; i++) { 
            if (desiredFolder === data.folders[i].name) { 
                folderIndex = i;
            }
        }
        if (folderIndex === -1) {
            throw ('Error - Invalid Folder Name');
        }
        return folderIndex;
    }

    /**
     * @param data dictionary object created by logInCreateData
     * @param desiredFolder string containing folder to search in
     * @return int total count of folders associated with account
     */
    folderQuantity(data, desiredFolder) { 
        let result = 0;  
        let folderIndex = this.folderExists(data, desiredFolder);
        for(let j = 0; j < data.messages.length; j++) {
            if (data.messages[j].folderId === data.folders[folderIndex].id) {
                result++;
            }
        }
        return result;
    }

    /**
     * @param data dictionary object created by logInCreateData
     * @param desiredFolder string containing folder to search in
     * @return array of all mail that exists in specified folder
     */
    folderMail(data, desiredFolder) {  
        let result = [];  
        let folderIndex = this.folderExists(data, desiredFolder);
        for(let j = 0; j < data.messages.length; j++) {
            if (data.messages[j].folderId === data.folders[folderIndex].id) {
                result.push(data.messages[j]);
            }
        }
        return result;
    }

    /**
     * @param data dictionary object created by logInCreateData 
     * @param to string specifying email address in to field
     * @param desiredFolder string containing folder to search in
     * @return int of all mail to specific email in desired search folder
     */
    toQuantity(data, to, desiredFolder) {  
        let result = 0;
        let folderIndex = this.folderExists(data, desiredFolder);
        for(let j = 0; j < data.messages.length; j++) { 
            if (data.messages[j].folderId === data.folders[folderIndex].id) { 
                for (let k = 0; k < data.messages[j].to.length; k++) {
                    if (data.messages[j].to[k].toUpperCase().trim() === to.toUpperCase().trim()) { 
                        result++;
                    }
                }
            }
        }
        return result;
    }

    /**
     * @param data dictionary object created by logInCreateData 
     * @param to string specifying email address in to field
     * @param desiredFolder string containing folder to search in
     * @return array of all mail objects to specific email in desired search folder
     */
    toMail(data, to, desiredFolder) {  
        let result = [];
        let folderIndex = this.folderExists(data, desiredFolder);
        for(let j = 0; j < data.messages.length; j++) { 
            if (data.messages[j].folderId === data.folders[folderIndex].id) { 
                for (let k = 0; k < data.messages[j].to.length; k++) {
                    if (data.messages[j].to[k].toUpperCase().trim() === to.toUpperCase().trim()) {
                        result.push(data.messages[j]);
                    }
                }
            }
        }
        return result;
    }

    /**
     * @param data dictionary object created by logInCreateData 
     * @param from string specifying email address in from field
     * @param desiredFolder string containing folder to search in
     * @return int of all mail from specific email in desired search folder
     */
    fromQuantity(data, from, desiredFolder) {  
        let result = 0;
        let folderIndex = this.folderExists(data, desiredFolder); 
        for(let j = 0; j < data.messages.length; j++) { 
            if (data.messages[j].folderId === data.folders[folderIndex].id) { 
                if (data.messages[j].from.toUpperCase().trim() === from.toUpperCase().trim()) {
                    result++;
                }
            }
        }
        return result;
    }

    /**
     * @param data dictionary object created by logInCreateData 
     * @param from string specifying email address in from field
     * @param desiredFolder string containing folder to search in
     * @return array of all mail from specific email in desired search folder
     */
    fromMail(data, from, desiredFolder) {  
        let result = [];
        let folderIndex = this.folderExists(data, desiredFolder);
        for(let j = 0; j < data.messages.length; j++) { 
            if (data.messages[j].folderId === data.folders[folderIndex].id) { 
                if (data.messages[j].from.toUpperCase().trim() === from.toUpperCase().trim()) {
                    result.push(data.messages[j]);
                }
            }
        }
        return result;
    }

    /**
     * @param data dictionary object created by logInCreateData 
     * @param subject string specifying subject match; is not case or empty space sensitive
     * @param desiredFolder string containing folder to search in
     * @return int of all mail with specific subject name in desired search folder
     */
    subjectQuantity(data, subject, desiredFolder) {  
        let result = 0;
        let folderIndex = this.folderExists(data, desiredFolder);
        for(let j = 0; j < data.messages.length; j++) { 
            if (data.messages[j].folderId === data.folders[folderIndex].id) { 
                if (data.messages[j].subject.toUpperCase().trim() === subject.toUpperCase().trim()) {
                    result++;
                }
            }
        }
        return result;
    }

    /**
     * @param data dictionary object created by logInCreateData 
     * @param subject string specifying subject match; is not case or empty space sensitive
     * @param desiredFolder string containing folder to search in
     * @return array of all mail with specific subject name in desired search folder
     */
    subjectMail(data, subject, desiredFolder) {  
        let result = [];
        let folderIndex = this.folderExists(data, desiredFolder);
        for(let j = 0; j < data.messages.length; j++) { 
            if (data.messages[j].folderId === data.folders[folderIndex].id) { 
                if (data.messages[j].subject.toUpperCase().trim() === subject.toUpperCase().trim()) {
                    result.push(data.messages[i]);
                }
            }
        }
        return result;
    }

    /**
     * @param data dictionary object created by logInCreateData 
     * @param startDate string "YYYY-MM-DD"
     * @param startTime string "HH:MM:SS"
     * @param finishDate string "YYYY-MM-DD"
     * @param finishTime string "HH:MM:SS"
     * @param desiredFolder string containing folder to search in
     * @return int total of emails received between desired time frames in desired search folder
     */
    dateTimeQuantity(data, startDate, startTime, finishDate, finishTime, desiredFolder) { 
        if (!startDate || !startTime || !finishDate || !finishTime) {
            throw ('Error - Invalid Date-Time Range');
        }
        let next = true;
        let nexter = 0;
        let question, middledate;
        let start = `${startDate}T${startTime}Z`.split(/[-T:Z]/);
        let finish = `${finishDate}T${finishTime}Z`.split(/[-T:Z]/);

        let folderIndex = this.folderExists(data, desiredFolder);
        for(let z = 0; z < data.messages.length; z++) { 
            if (data.messages[z].folderId === data.folders[folderIndex].id) { 
                next = false;
                question = data.messages[z].timestamp.split(/[-T:Z]/);
                middledate = false;
                for (let x = 0; x < 3; x++) {
                    if(Number(question[x]) >= Number(start[x]) && Number(question[x]) <= Number(finish[x])){ 
                        middledate = true;
                    }
                }
                if (middledate = true) {
                    for (let j = 0; j < 3; j++){
                        if(Number(question[j]) > Number(start[j]) && Number(question[j]) < Number(finish[j])){ 
                            next = true;
                        }
                        else if(Number(question[j]) === Number(start[j]) && Number(question[j]) === Number(finish[j])){ 
                            if(Number(question[3]) < Number(finish[3])){
                                next = true;
                                break;
                            }
                            else if(Number(question[3]) === Number(finish[3])){
                                if(Number(question[4]) < Number(finish[4])){
                                    next = true;
                                    break;
                                }
                                else if(Number(question[4]) === Number(finish[4])){
                                    if(Number(question[5]) < Number(finish[5])){
                                        next = true;
                                        break;
                                    }
                                }
                            }
                        }
                        else if (Number(question[j]) === Number(start[j]) && Number(question[j]) < Number(finish[j])){
                            if(Number(question[3]) < Number(finish[3])){
                                next = true;
                                break;
                            }
                            else if(Number(question[3]) === Number(finish[3])){
                                if(Number(question[4]) < Number(finish[4])){
                                    next = true;
                                    break;
                                }
                                else if(Number(question[4]) === Number(finish[4])){
                                    if(Number(question[5]) < Number(finish[5])){
                                        next = true;
                                        break;
                                    }
                                }
                            }
                        }
                        else if (Number(question[j]) > Number(start[j]) && Number(question[j]) === Number(finish[j])){
                            if(Number(question[3]) < Number(finish[3])){
                                next = true;
                                break;
                            }
                            else if(Number(question[3]) === Number(finish[3])){
                                if(Number(question[4]) < Number(finish[4])){
                                    next = true;
                                    break;
                                }
                                else if(Number(question[4]) === Number(finish[4])){
                                    if(Number(question[5]) < Number(finish[5])){
                                        next = true;
                                        break;
                                    }
                                }
                            }
                        }
                    }
                }
                if(next){
                    nexter++;
                } 
            }
        }
        return nexter;
    }

    /**
     * @param data dictionary object created by logInCreateData 
     * @param startDate string "YYYY-MM-DD"
     * @param startTime string "HH:MM:SS"
     * @param finishDate string "YYYY-MM-DD"
     * @param finishTime string "HH:MM:SS"
     * @param desiredFolder string containing folder to search in
     * @return array of emails received between desired time frames in desired search folder
     */
    dateTimeMail(data, startDate, startTime, finishDate, finishTime, desiredFolder) {  
        if (!startDate || !startTime || !finishDate || !finishTime) {
            throw ('Error - Invalid Date-Time Range');
        }
        let next = true;
        let result = [];
        let question, middledate;
        let start = `${startDate}T${startTime}Z`.split(/[-T:Z]/);
        let finish = `${finishDate}T${finishTime}Z`.split(/[-T:Z]/);
        let folderIndex = this.folderExists(data, desiredFolder);
        for(let z = 0; z < data.messages.length; z++) { 
            if (data.messages[z].folderId === data.folders[folderIndex].id) { 
                next = false;
                question = data.messages[z].timestamp.split(/[-T:Z]/);
                middledate = false;
                for (let x = 0; x < 3; x++) {
                    if(Number(question[x]) >= Number(start[x]) && Number(question[x]) <= Number(finish[x])){ 
                        middledate = true;
                    }
                }
                if (middledate = true) {
                    for (let j = 0; j < 3; j++){
                        if(Number(question[j]) > Number(start[j]) && Number(question[j]) < Number(finish[j])){ 
                            next = true;
                        }
                        else if(Number(question[j]) === Number(start[j]) && Number(question[j]) === Number(finish[j])){ 
                            if(Number(question[3]) < Number(finish[3])){
                                next = true;
                                break;
                            }
                            else if(Number(question[3]) === Number(finish[3])){
                                if(Number(question[4]) < Number(finish[4])){
                                    next = true;
                                    break;
                                }
                                else if(Number(question[4]) === Number(finish[4])){
                                    if(Number(question[5]) < Number(finish[5])){
                                        next = true;
                                        break;
                                    }
                                }
                            }
                        }
                        else if (Number(question[j]) === Number(start[j]) && Number(question[j]) < Number(finish[j])){
                            if(Number(question[3]) < Number(finish[3])){
                                next = true;
                                break;
                            }
                            else if(Number(question[3]) === Number(finish[3])){
                                if(Number(question[4]) < Number(finish[4])){
                                    next = true;
                                    break;
                                }
                                else if(Number(question[4]) === Number(finish[4])){
                                    if(Number(question[5]) < Number(finish[5])){
                                        next = true;
                                        break;
                                    }
                                }
                            }
                        }
                        else if (Number(question[j]) > Number(start[j]) && Number(question[j]) === Number(finish[j])){
                            if(Number(question[3]) < Number(finish[3])){
                                next = true;
                                break;
                            }
                            else if(Number(question[3]) === Number(finish[3])){
                                if(Number(question[4]) < Number(finish[4])){
                                    next = true;
                                    break;
                                }
                                else if(Number(question[4]) === Number(finish[4])){
                                    if(Number(question[5]) < Number(finish[5])){
                                        next = true;
                                        break;
                                    }
                                }
                            }
                        }
                    }
                }
                if(next){
                    result.push(data.messages[z]);
                } 
            }
        }
        return result;
    }

    /**
     * @param data dictionary object created by logInCreateData 
     * @param str string containing search string
     * @param desiredFolder string containing folder to search in
     * @return array containing all emails with bodies that have the desired search string in them
     */
    bodyContainsString(data, str, desiredFolder) {  
        let result = [];
        let folderIndex = this.folderExists(data, desiredFolder);
        for(let j = 0; j < data.messages.length; j++) { 
            if (data.messages[j].folderId === data.folders[folderIndex].id) { 
                if (data.messages[j].body.toUpperCase().trim().indexOf(str.toUpperCase().trim()) > -1) {
                        result.push(data.messages[j]);                                               
                }
            }
        }
        return result;
    }
}

module.exports = Fetch;
