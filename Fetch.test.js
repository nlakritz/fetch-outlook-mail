const Fetch = require('./Fetch.js');
const sampleData = require('./sampleMail.json');

describe(`Fetch test suite`, () => {
    describe(`test constructor`, () => {
        // test 1
        it(`should throw error if no username or password provided`, () => {
            expect(() => {
                let fetcher = new Fetch('sample.address@foo.com', '');
            }).toThrow('Invalid object initialization - must provide username and password');
            expect(() => {
                let fetcher2 = new Fetch('', 'password123!');
            }).toThrow('Invalid object initialization - must provide username and password');
            expect(() => {
                let fetcher3 = new Fetch('', '');
            }).toThrow('Invalid object initialization - must provide username and password');
        });
        // test 2
        it(`should initialize object`, () => {
            let fetcher = new Fetch('sample.address@foo.com', 'asdf');
            expect(fetcher._username).toEqual('sample.address@foo.com');
            expect(fetcher._password).toEqual('asdf');
            expect(fetcher._tokenOptions).not.toEqual(null);   
        });
    });
    describe(`test getter functions`, () => {
        let fetcher;
        beforeEach(() => {
            fetcher = new Fetch('sample.address@foo.com', 'pass123');
        });
        afterEach(() => {
            fetcher = null;
        });
        // test 3
        it(`should return username`, () => {
            let username = fetcher.username;
            expect(username).toEqual('sample.address@foo.com');
        });
        // test 4
        it(`should return password`, () => {
            let pass = fetcher.password;
            expect(pass).toEqual('pass123');
        });
        // test 5
        it(`should return token options object`, () => {
            let token = {
                url: 'https://login.windows.net/common/oauth2/token', 
                headers: { 'content-type': 'application/x-www-form-urlencoded' },
                json: true };
            expect(fetcher.tokenOptions.url).toEqual(token.url);
            expect(fetcher.tokenOptions.headers).toEqual(token.headers);
            expect(fetcher.tokenOptions.json).toEqual(token.json);
        });
    });
    describe(`test setter functions`, () => {
        let fetcher;
        beforeEach(() => {
            fetcher = new Fetch('sample.address@foo.com', 'pass123');
        });
        afterEach(() => {
            fetcher = null;
        });
        // test 6 
        it(`should set username`, () => {
            let user = 'sampleuser';
            fetcher.username = user;
            expect(fetcher.username).toEqual(user);
        });
        // test 7
        it(`should set password`, () => {
            let pass = 'samplepass';
            fetcher.password = pass;
            expect(fetcher.password).toEqual(pass);
        });
    });

    let fetcher;
    let inboxId = "AQMkAGYzNmEwMAA5Mi01ZGZjLTRkZTQtODM2Yi0wMjY5NDc2ODdhYjYALgAAA0ej4msDk8tLiWUK3LkhzWQBAD6DO4zu5CJOvkSSgsE4pUkAAAIBDAAAAA==";
    beforeEach(() => {
        fetcher = new Fetch('sample.address@foo.com', 'pass123');
    });
    afterEach(() => {
        fetcher = null;
    });
    describe(`folderList`, () => {
        // test 8
        it(`should contain all existing folders`, () => {
            let folders = fetcher.folderList(sampleData);
            expect(folders[0].name).toEqual('Archive');
            expect(folders[1].name).toEqual('Conversation History');
            expect(folders[2].name).toEqual('Deleted Items');
            expect(folders[3].name).toEqual('Drafts');
            expect(folders[4].name).toEqual('Inbox');
            expect(folders[5].name).toEqual('Junk Email');expect(folders[6].name).toEqual('Outbox');
            expect(folders[7].name).toEqual('RSS Subscriptions');
            expect(folders[8].name).toEqual('Sent Items');
            expect(folders[9].name).toEqual('Sync Issues');
        });
        // test 9
        it(`should not include non-existing folders`, () => {
            let folders = fetcher.folderList(sampleData);
            expect(folders[0].name).not.toEqual('Invalid Folder');
            expect(() => {
                expect(folders[10].name);
            }).toThrowError();
        });
        // test 10
        it(`should show error when sample data is invalid`, () => {
            expect(() => {
                folders = fetcher.folderList(null);
            }).toThrowError();
        });
        // test 11
        it(`should have a name and id for each folder`, () => {
            let folders = fetcher.folderList(sampleData);
            expect(folders[0].name).toBeTruthy();
            expect(folders[0].id).toBeTruthy();
        });
    });
    describe(`profileInfo`, () => {
        // test 12
        it(`should contain all existing profile data`, () => {
            let profile = fetcher.profileInfo(sampleData);
            expect(profile['Display Name']).toEqual('John Smith');
            expect(profile['Email Address']).toEqual('sample.address@foo.com');
        });
        // test 13
        it(`should throw error when sample data is invalid`, () => {
            expect(() => {
                let profile = fetcher.profileInfo(null);
            }).toThrowError();
        });
    });
    describe(`totalQuantity`, () => {
        // test 14
        it(`should return total number of emails across all folders`, () => {
            let total = fetcher.totalQuantity(sampleData);
            expect(total).toEqual(sampleData.messages.length);
        });
    });
    describe(`allMail`, () => {
        // test 15
        it(`should return contents of emails across all folders`, () => {
            let total = fetcher.allMail(sampleData);
            expect(total).toEqual(sampleData.messages);
            expect(total.length).toEqual(sampleData.messages.length);
        });
    });
    describe(`folderExists`, () => {
        // test 16
        it(`should return zero or greater folder index if folder exists`, () => {
            let finder = fetcher.folderExists(sampleData, "Inbox");
            expect(finder).toBeGreaterThan(-1);
            expect(finder).toBeLessThan(sampleData.folders.length-1)
        });
        // test 17
        it(`should throw error if folder does not exist`, () => {
            expect(() => {
                let finder = fetcher.folderExists(sampleData, "Inbex");
            }).toThrowError('Error - Invalid Folder Name');
        });
    });
    describe(`folderQuantity`, () => {
        // test 18
        it(`should return number of messages in valid folder`, () => {
            let quantity = fetcher.folderQuantity(sampleData, "Inbox");
            let nexter = 0;
            for(let i = 0; i < sampleData.messages.length; i++) {
                if(sampleData.messages[i].folderId === inboxId) {
                    nexter++;
                }
            }
            expect(quantity).toEqual(nexter);                                        
        });
        // test 19
        it(`should throw error if folder does not exist`, () => {
            expect(() => {
                let quantity = fetcher.folderQuantity(sampleData, "Inbex");
            }).toThrowError('Error - Invalid Folder Name');
        });
    });
    describe(`folderMail`, () => {
        // test 20
        it(`should return contents of messages in valid folder`, () => {
            let actual = fetcher.folderMail(sampleData, "Inbox");
            let expected = [];
            for(let i = 0; i < sampleData.messages.length; i++) {
                if(sampleData.messages[i].folderId === inboxId) {
                    expected.push(sampleData.messages[i]);
                }
            }
            expect(actual).toEqual(expected);   
        });
    });
    describe(`toQuantity`, () => {
        // test 21
        it(`should return number of messages with 'to' field in valid folder`, () => {
            let quantity = fetcher.toQuantity(sampleData, "sample2@foo.no", "Inbox");
            let nexter = 0; 
            for(let i = 0; i < sampleData.messages.length; i++) {
                if(sampleData.messages[i].folderId === inboxId) {
                    for(let j = 0; j < sampleData.messages[i].to.length; j++) {
                        if(sampleData.messages[i].to[j].toUpperCase().trim() === "sample2@foo.no".toUpperCase().trim()){
                            nexter++
                        }
                    } 
                }
            }
            expect(quantity).toEqual(nexter);                                             
        });
    });
    describe(`toMail`, () => {
        // test 22
        it(`should return contents of messages with 'to' field in valid folder`, () => {
            let actual = fetcher.toMail(sampleData, "sample3@foo.no", "Inbox");
            let expected = []; 
            for(let i = 0; i < sampleData.messages.length; i++) {
                if(sampleData.messages[i].folderId === inboxId) {
                    for(let j = 0; j < sampleData.messages[i].to.length; j++) {
                        if(sampleData.messages[i].to[j].toUpperCase().trim() === "sample3@foo.no".toUpperCase().trim()){
                            expected.push(sampleData.messages[i]);
                        }
                    } 
                }
            }
            expect(actual).toEqual(expected);                                  
        });
    });
    describe(`fromQuantity`, () => {
        // test 23
        it(`should return number of messages with 'from' field in valid folder`, () => {
            let quantity = fetcher.fromQuantity(sampleData, "sample3@foo.no", "Inbox");
            let nexter = 0;
            for(let i = 0; i < sampleData.messages.length; i++) {
                if(sampleData.messages[i].folderId === inboxId) {
                    if(sampleData.messages[i].from.toUpperCase().trim() === "sample3@foo.no".toUpperCase().trim()){
                        nexter++;
                    }
                }
            }
            expect(quantity).toEqual(nexter);                                          
        });
    });
    describe(`fromMail`, () => {
        // test 24
        it(`should return contents of messages with 'from' field in valid folder`, () => {
            let actual = fetcher.fromMail(sampleData, "sample3@foo.no", "Inbox");
            let expected = [];
            for(let i = 0; i < sampleData.messages.length; i++) {
                if(sampleData.messages[i].folderId === inboxId) {
                    if(sampleData.messages[i].from.toUpperCase().trim() === "sample3@foo.no".toUpperCase().trim()){
                        expected.push(sampleData.messages[i]);;
                    }
                }
            }
            expect(actual).toEqual(expected);                                             
        });
    });
    describe(`subjectQuantity`, () => {
        // test 25
        it(`should return number of messages with 'subject' field in valid folder`, () => {
            let quantity = fetcher.subjectQuantity(sampleData, "Testing", "Inbox");
            let nexter = 0;
            for(let i = 0; i < sampleData.messages.length; i++) {
                if(sampleData.messages[i].folderId === inboxId) {
                    if(sampleData.messages[i].subject.toUpperCase().trim() === "Testing".toUpperCase().trim()){
                        nexter++;
                    }
                }
            }
            expect(quantity).toEqual(nexter);                                           
        });
    });
    describe(`subjectMail`, () => {
        // test 26
        it(`should return contents of messages with 'subject' field in valid folder`, () => {
            let actual = fetcher.subjectMail(sampleData, "Testing", "Inbox");
            let expected = [];
            for(let i = 0; i < sampleData.messages.length; i++) {
                if(sampleData.messages[i].folderId === inboxId) {
                    if(sampleData.messages[i].from.toUpperCase().trim() === "Testing".toUpperCase().trim()){
                        expected.push(sampleData.messages[i]);;
                    }
                }
            }
            expect(actual).toEqual(expected);                                           
        });
    });
    describe(`dateTimeQuantity`, () => {
        // test 27
        it(`should return number of messages within valid date-time range in valid folder`, () => {
            let quantity = fetcher.dateTimeQuantity(sampleData, "2018-07-10", "00:07:35", "2018-07-12", "00:08:00", "Inbox");
            expect(quantity).toEqual(2);
        });
        // test 28
        it(`should throw error if entire date-time range is not provided`, () => {
            expect(() => {
                let quantity = fetcher.dateTimeQuantity(sampleData, "Inbox");
            }).toThrowError('Error - Invalid Date-Time Range');
        });
    });
    describe(`dateTimeMail`, () => {
        // test 29
        it(`should return contents of messages within valid date-time range in valid folder`, () => {
            let actual = fetcher.dateTimeMail(sampleData, "2018-07-10", "00:07:35", "2018-07-12", "00:08:00", "Inbox");
            expected = [];
            expected.push(sampleData.messages[3]);
            expected.push(sampleData.messages[4]);
            expect(actual).toEqual(expected);                               
        });
        // test 30
        it(`should throw error if entire date-time range is not provided`, () => {
            expect(() => {
                let actual = fetcher.dateTimeMail(sampleData, '2018-07-10', '00:06:00', "Inbox");
            }).toThrowError('Error - Invalid Date-Time Range');
        });
    });
    describe(`bodyContainsString`, () => {
        // test 31
        it(`should return contents of messages containing body in valid folder`, () => {
            let actual = fetcher.bodyContainsString(sampleData, "fasdf", "Inbox");    
            let expected = [].concat(sampleData.messages[4], sampleData.messages[8]);
            expect(actual).toEqual(expected);    
        });
    });
});
