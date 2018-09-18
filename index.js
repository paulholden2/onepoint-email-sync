const fs = require('fs');
const async = require('async');
const readline = require('readline');
const { google } = require('googleapis');
const dotenv = require('dotenv');
const OnePoint = require('onepoint-node-sdk');

dotenv.config();

// If modifying these scopes, delete token.json.
const SCOPES = [
  'https://www.googleapis.com/auth/admin.directory.user',
  'https://www.googleapis.com/auth/admin.directory.group'
];

const TOKEN_PATH = 'token.json';

const EVERYONE_EMAIL = 'everyonetest@stria.com';
const EVERYONE_NAME = 'Everyone (Test)';

// Load client secrets from a local file.
fs.readFile('credentials.json', (err, content) => {
  if (err) {
    return console.error('Error loading client secret file', err);
  }

  // Authorize a client with the loaded credentials, then call the
  // Directory API.
  authorize(JSON.parse(content), onePointSync);
});

/**
 * Create an OAuth2 client with the given credentials, and then execute the
 * given callback function.
 *
 * @param {Object} credentials The authorization client credentials.
 * @param {function} callback The callback to call with the authorized client.
 */
function authorize(credentials, callback) {
  const { client_secret, client_id, redirect_uris } = credentials.installed;
  const oauth2Client = new google.auth.OAuth2(client_id, client_secret, redirect_uris[0]);

  // Check if we have previously stored a token.
  fs.readFile(TOKEN_PATH, (err, token) => {
    if (err) {
      return getNewToken(oauth2Client, callback);
    }

    oauth2Client.credentials = JSON.parse(token);

    callback(oauth2Client);
  });
}

/**
 * Get and store new token after prompting for user authorization, and then
 * execute the given callback with the authorized OAuth2 client.
 *
 * @param {google.auth.OAuth2} oauth2Client The OAuth2 client to get token for.
 * @param {getEventsCallback} callback The callback to call with the authorized
 *     client.
 */
function getNewToken(oauth2Client, callback) {
  const authUrl = oauth2Client.generateAuthUrl({
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
    oauth2Client.getToken(code, (err, token) => {
      if (err) {
        return console.error('Error retrieving access token', err);
      }

      oauth2Client.credentials = token;
      storeToken(token);

      callback(oauth2Client);
    });
  });
}

/**
 * Store token to disk be used in later program executions.
 *
 * @param {Object} token The token to store to disk.
 */
function storeToken(token) {
  fs.writeFile(TOKEN_PATH, JSON.stringify(token), (err) => {
    if (err) {
      return console.warn(`Token not stored to ${TOKEN_PATH}`, err);
    }

    console.log(`Token stored to ${TOKEN_PATH}`);
  });
}

function paced(callback) {
  return () => {
    setTimeout(callback, 250);
  }
}

/**
 * Syncs OnePoint emails with Stria emails into an everyone@stria.com group
 *
 * @param {google.auth.OAuth2} auth An authorized OAuth2 client.
 */
function onePointSync(auth) {
  const service = google.admin({
    version: 'directory_v1',
    auth: auth
  });

  var onePoint;
  var emailMembers;
  var employeeList;

  async.waterfall([
    /**
     * Connect to OnePoint
     */
    (callback) => {
      var cl = new OnePoint({
        username: process.env.ONEPOINT_USERNAME,
        password: process.env.ONEPOINT_PASSWORD,
        companyShortName: process.env.ONEPOINT_COMPANY_SHORT_NAME,
        apiKey: process.env.ONEPOINT_API_KEY
      });

      cl.connect((err) => {
        if (err) {
          return callback(err);
        }

        onePoint = cl;

        callback();
      });
    },
    (callback) => {
      onePoint.runReport({
        where: {
          savedName: 'Employee Emails'
        }
      }, (err, reports) => {
        if (err) {
          return callback(err);
        }

        employeeList = reports[0].results;

        callback();
      });
    },
    /**
     * Create the everyone@stria.com group
     */
    (callback) => {
      service.groups.insert({
        resource: {
          email: EVERYONE_EMAIL,
          name: EVERYONE_NAME,
          description: 'Contains all Stria email addresses and those of employees without a stria.com email address'
        }
      }, (err, res) => {
        if (err) {
          // Exists
          return callback();
        }

        callback();
      });
    },
    /**
     * Create a list of all employees already in the everyone@stria.com group
     */
    (callback) => {
      var members = [];
      var nextPageToken = false;

      var short = callback;

      async.doWhilst((callback) => {
        service.members.list({
          groupKey: EVERYONE_EMAIL,
          maxResults: 200,
          pageToken: nextPageToken || undefined
        }, (err, res) => {
          if (err) {
            // No members
            return short(null, members);
          }

          nextPageToken = res.data.nextPageToken || false;
          members = members.concat(res.data.members.filter((m) => {
            return m.status === 'ACTIVE' && m.type === 'USER'
          }));

          callback();
        });
      }, () => nextPageToken, (err) => {
        if (err) {
          return callback(err);
        }

        callback(null, members);
      });
    },
    (members, callback) => {
      emailMembers = members;

      callback();
    },
    /**
     * Remove all members of the everyone@stria.com email group that aren't
     * in the active employees report
     */
    (callback) => {
      var employeeEmails = employeeList.map(e => e['"Email"'].toLowerCase());

      async.each(emailMembers, (member, callback) => {
        if (employeeEmails.indexOf(member.email.toLowerCase()) < 0) {
          console.log('Removing terminated employee: ' + member.email);
          service.members.delete({
            groupKey: EVERYONE_EMAIL,
            memberKey: member.email
          }, paced(callback));
        } else {
          paced(callback)();
        }
      }, callback);
    },
    /**
     * Add all employees not already in the list.
     */
    (callback) => {
      var employeeEmails = emailMembers.map(e => e.email.toLowerCase());

      async.eachSeries(employeeList, (employee, callback) => {
        var employeeEmail = employee['"Email"'];

        if (employeeEmails.indexOf(employeeEmail.toLowerCase()) < 0) {
          console.log('Adding active employee: ' + employeeEmail);

          service.members.insert({
            groupKey: EVERYONE_EMAIL,
            resource: {
              email: employeeEmail,
              role: 'MEMBER'
            }
          }, paced(() => callback()));
        } else {
          paced(callback)();
        }
      }, callback);
    }
  ], (err) => {
    var code = 0;

    if (err) {
      code = 1;
      console.log(err);
    } else {
      console.log('Done');
    }

    process.exit(code);
  });
}
