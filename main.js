const fs = require('fs');
const readline = require('readline');
const { google } = require('googleapis');
const spreadId = '165i3dys_XEaq2lKcKiJL31IWIs5IaG7Q7yK3aASnQFQ';
const bot = require('node-rocketchat-bot');
const keys = require('./keys.json');

// If modifying these scopes, delete token.json.
const SCOPES = ['https://www.googleapis.com/auth/spreadsheets'];
// The file token.json stores the user's access and refresh tokens, and is
// created automatically when the authorization flow completes for the first
// time.
const TOKEN_PATH = 'token.json';

// Load client secrets from a local file.
fs.readFile('credentials.json', (err, content) => {
    if (err) return console.log('Error loading client secret file:', err);
    // Authorize a client with credentials, then call the Google Sheets API.
    //authorize(JSON.parse(content), listDevices);
    authorize(JSON.parse(content), main);
});

/**
 * Create an OAuth2 client with the given credentials, and then execute the
 * given callback function.
 * @param {Object} credentials The authorization client credentials.
 * @param {function} callback The callback to call with the authorized client.
 */
function authorize(credentials, callback) {
    const { client_secret, client_id, redirect_uris } = credentials.installed;
    const oAuth2Client = new google.auth.OAuth2(
        client_id, client_secret, redirect_uris[0]);

    // Check if we have previously stored a token.
    fs.readFile(TOKEN_PATH, (err, token) => {
        if (err) return getNewToken(oAuth2Client, callback);
        oAuth2Client.setCredentials(JSON.parse(token));
        callback(oAuth2Client);
    });
}

/**
 * Get and store new token after prompting for user authorization, and then
 * execute the given callback with the authorized OAuth2 client.
 * @param {google.auth.OAuth2} oAuth2Client The OAuth2 client to get token for.
 * @param {getEventsCallback} callback The callback for the authorized client.
 */
function getNewToken(oAuth2Client, callback) {
    const authUrl = oAuth2Client.generateAuthUrl({
        access_type: 'offline',
        scope: SCOPES,
    });
    console.log('Authorize this app by visiting this url:', authUrl);
    const rl = readline.createInterface({
        input: process.stdin,
        output: process.stdout,
    });
    rl.question('Enter the code from that page here: ', (code) => {
        rl.close();
        oAuth2Client.getToken(code, (err, token) => {
            if (err) return console.error('Error while trying to retrieve access token', err);
            oAuth2Client.setCredentials(token);
            // Store the token to disk for later program executions
            fs.writeFile(TOKEN_PATH, JSON.stringify(token), (err) => {
                if (err) return console.error(err);
                console.log('Token stored to', TOKEN_PATH);
            });
            callback(oAuth2Client);
        });
    });
}

/**
 * Prints the names and majors of students in a sample spreadsheet:
 * @see https://docs.google.com/spreadsheets/d/1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms/edit
 * @param {google.auth.OAuth2} auth The authenticated Google OAuth client.
 */
function listDevices(auth) {
    const sheets = google.sheets({ version: 'v4', auth });
    sheets.spreadsheets.values.get({
        spreadsheetId: spreadId,
        range: 'A2:A100',
    }, (err, res) => {
        if (err) return console.log('The API returned an error: ' + err);
        const rows = res.data.values;
        if (rows.length) {
            console.log('Devices:');
            // Print columns A and E, which correspond to indices 0 and 4.
            rows.map((row) => {
                console.log(`${row}`);
            });
        } else {
            console.log('No data found.');
        }
    });
}

function main(auth) {
    //console.log(words);

    bot({
        host: keys.host,
        username: keys.username,
        password: keys.password,
        // use ssl for https
        ssl: true,
        pretty: false,
        // join room(s)
        rooms: ['bots'],
        // when ready (e.log.info logs to console, can also console.log)
        onWake: async event => event.log.info(`${event.bot.username} ready`),
        // on message
        onMessage: async event => {
            if (event.flags.isMentioned) {
                const words = event.message.content.split(' ');
                const operation = words[1] ? words[1].toLowerCase() : ''
                event.log.info(`operation is "${operation}"`)
                processCommand(auth, words, event);
            }
        }
    });
}

function processCommand(auth, words, event) {
    switch (words[1].toLowerCase()) {
        case "help":
            event.respond("This bot is for checking in and checking out Devices \n\"Checkout LaptopId UserId\" checks out LaptopId to UserID\n\t(chromebook #XXXX is checked out to user #XXXX)" +
                "\n\"Checkin LaptopID\" checks in LaptopId \n\t(chromebook #XXXX has been checked in by user #XXXX)");
            break;
        case "checkout":
            if (words[2] == undefined && words[3] == undefined) {
                event.respond("Incorrect Input, please try again or use help");
            }
            else {
                checkOut(auth, words[2], words[3], event);
            }

            break;
        case "checkin":
            if (words[2] == undefined) {
                event.respond("Incorrect Input, please try again or use help");
            }
            else {
                checkIn(auth, words[2], event);
            }
            break;
        default:
            event.respond("Incorrect Input, please try again or use help");
            break;
    }
}

//used to get input from multiple fields. Needed for the scan gun to work
//scan gun inputs text then hits enter key. Therefore needs to have multiple lines of questioning
function getCheckOutInfo(auth, event) {
}


//checks out a laptop to a user
async function checkOut(auth, device, user, event) {
    event.respond("Checking out " + device + " to User " + user);

    const sheets = google.sheets({ version: 'v4', auth });

    const opt = {
        spreadsheetId: spreadId, //spreadsheet id
        range: 'A2:C100'    //value range we are looking at, we need to check E2:E100 to see if there is a clock on time
    };

    let data = await sheets.spreadsheets.values.get(opt);
    dataArray = data.data.values;

    let found = false;
    let position = 2;

    dataArray.forEach(element => {
        if (element[0] == device) {
            found = true;
            //if the device is in, therefore ready to be checked out
            if (element[1] == "in") {
                //need to change the element to have the users ID attached
                updateDeviceOut(auth, position, user);
            }
            else {
                event.respond("That device is listed as already checked out. Please check the log and verify this.");
                console.log("That device is listed as already checked out. \nPlease check the log and verify this.");
            }
        }
        position += 1;
    });

    if (found == false) {
        event.respond("No device " + device + " was found, please make sure it was inputted correctly and that the device exists in the spreadsheet.");
        console.log("No device " + device + " was found, please make sure it was inputted correctly and that the device exists in the spreadsheet.");
    }
}

//checks in a laptop
async function checkIn(auth, device, event) {
    event.respond("Checking in " + device);

    const sheets = google.sheets({ version: 'v4', auth });

    const opt = {
        spreadsheetId: spreadId, //spreadsheet id
        range: 'A2:C100'    //value range we are looking at, we need to check E2:E100 to see if there is a clock on time
    };

    let data = await sheets.spreadsheets.values.get(opt);
    dataArray = data.data.values;

    let position = 2;
    let found = false;

    dataArray.forEach(element => {
        if (element[0] == device) {
            found = true;
            //if the device is out, therefore ready to be checked in
            if (element[1] == "out") {
                //need to change the element to have the users ID attached
                updateDeviceIn(auth, position);
            }
            //otherwise the laptop was never checked out
            else {
                event.respond("That device is listed as already checked in. Please check the log and verify this.");
                console.log("That device is listed as already checked in. \nPlease check the log and verify this.");
            }
        }
        position += 1;
    });

    if (found == false) {
        event.respond("No device " + device + " was found, please make sure it was inputted correctly and that the device exists in the spreadsheet.");
        console.log("No device " + device + " was found, please make sure it was inputted correctly and that the device exists in the spreadsheet.");
    }
}

//updates the device to have the users ID attached to it and the state flipped to out
async function updateDeviceOut(auth, pos, user) {
    var date = new Date();

    var fullDate = date.getMonth() + 1 + "/" + date.getDate() + "/" + date.getFullYear();
    var hour = date.getHours();
    var newRange = 'LaptopLog!A' + pos;


    const sheets = google.sheets({ version: 'v4', auth });

    const opt = {
        spreadsheetId: spreadId, //spreadsheet id
        range: 'A2:C100'    //value range we are looking at, we need to check E2:E100 to see if there is a clock on time
    };


    let data = await sheets.spreadsheets.values.get(opt);
    dataArray = data.data.values;

    vals = {
        "range": "LaptopLog!A" + pos,
        "majorDimension": "ROWS",
        "values": [
            [null, "out", user],
        ],
    };

    const updateOptions = {
        spreadsheetId: spreadId,
        range: newRange,
        valueInputOption: 'USER_ENTERED',
        resource: vals,
    };



    let res = await sheets.spreadsheets.values.update(updateOptions);
}

//updates the device to remove users ID and set the state to In
async function updateDeviceIn(auth, pos) {
    var date = new Date();

    var fullDate = date.getMonth() + 1 + "/" + date.getDate() + "/" + date.getFullYear();
    var hour = date.getHours();
    var newRange = 'LaptopLog!A' + pos;


    const sheets = google.sheets({ version: 'v4', auth });

    const opt = {
        spreadsheetId: spreadId, //spreadsheet id
        range: 'A2:C100'    //value range we are looking at, we need to check E2:E100 to see if there is a clock on time
    };


    let data = await sheets.spreadsheets.values.get(opt);
    dataArray = data.data.values;

    vals = {
        "range": "LaptopLog!A" + pos,
        "majorDimension": "ROWS",
        "values": [
            [null, "in", ""],
        ],
    };

    const updateOptions = {
        spreadsheetId: spreadId,
        range: newRange,
        valueInputOption: 'USER_ENTERED',
        resource: vals,
    };



    let res = await sheets.spreadsheets.values.update(updateOptions);

}

//returns true if laptop is checked out and false if not checked out
//column B keeps track of state
async function isChecked(auth, device) {
    const sheets = google.sheets({ version: 'v4', auth });

    const opt = {
        spreadsheetId: spreadId, //spreadsheet id
        range: 'A2:C100'    //value range we are looking at, we need to check E2:E100 to see if there is a clock on time
    };

    let data = await sheets.spreadsheets.values.get(opt);
    dataArray = data.data.values;

    let position = 0;
    let checked = 0;

    dataArray.forEach(element => {
        if (element[0] == device) {
            checked = position;
        }
        position += 1;
    });

    if (dataArray[checked][1] == "in") {
        console.log(dataArray[checked]);
        return false;
    }
    else {
        console.log(dataArray[checked]);
        return true;
    }
    //return checked;
}


