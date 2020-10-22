import { AzureFunction, Context, HttpRequest } from "@azure/functions"
import axios, { AxiosRequestConfig } from 'axios';
import qs = require('qs');


const SUPPORT_EMAIL = "thomas@balderdash.ch"

const APP_ID = "";
const APP_SECERET = "";
const TENANT_ID = "";

const TWILIO_ACCOUNT_SID = '';
const TWILIO_AUTH_TOKEN = '';
const TWILIO_NUMBER = '';

const TOKEN_ENDPOINT = 'https://login.microsoftonline.com/' + TENANT_ID + '/oauth2/v2.0/token';
const MS_GRAPH_SCOPE = 'https://graph.microsoft.com/.default';
const MS_GRAPH_ENDPOINT = 'https://graph.microsoft.com/v1.0/';
const MS_GRAPH_ENDPOINT_BETA = 'https://graph.microsoft.com/beta/';

const twilio = require('twilio')(TWILIO_ACCOUNT_SID, TWILIO_AUTH_TOKEN);


const httpTrigger: AzureFunction = async function (context: Context, req: HttpRequest): Promise<void> {
    context.log('HTTP trigger function processed a request.');

    axios.defaults.headers.post['Content-Type'] = 'application/x-www-form-urlencoded';

    let token = await getToken();    
    let riskyusers = await getRiskyUsers(token);
    let userinfos = await getRiskyUserProfiles(token, riskyusers);
    for(let i in userinfos) {
        resetPassword(token, userinfos[i]);
        //await sendSMS(userinfos[i].mobilePhone, userinfos[i].userPrincipalName);
        /*for (let mail of userinfos[i].otherMails) {
            await sendMail(token, userinfos[i].userPrincipalName, mail);
        }
        await sendSupportMail(token, userinfos[i], riskyusers[i])*/
    }
    
    context.res = {
        // status: 200, /* Defaults to 200 */
        body: userinfos
    };

};

export default httpTrigger;

class RiskyUser {
    id: string;
    isDeleted: boolean;
    isProcessing: false;
    riskLevel: string;
    riskState: string;
    riskDetail: string;
    riskLastUpdatedDateTime: string;
    userDisplayName: string;
    userPrincipalName: string;
}

class RiskyUserProfile {
    displayName: string;
    userPrincipalName: string;
    mobilePhone: string;
    otherMails: string[];
}

async function getToken(): Promise<string> {
    const postData = {
        client_id: APP_ID,
        scope: MS_GRAPH_SCOPE,
        client_secret: APP_SECERET,
        grant_type: 'client_credentials'
    };

    return await axios
        .post(TOKEN_ENDPOINT, qs.stringify(postData))
        .then(response => {
            // console.log(response.data);
            return response.data.access_token;
        })
        .catch(error => {
            console.log(error);
        });
}

async function getRiskyUsers(token:string): Promise<RiskyUser[]> {
    let config: AxiosRequestConfig = {
        method: 'get',
        url: MS_GRAPH_ENDPOINT + 'identityProtection/riskyUsers',
        headers: {
          'Authorization': 'Bearer ' + token //the token is a variable which holds the token
        }
    }
    
    return await axios(config)
        .then(response => {
            // console.log(response.data);
            return response.data.value;
        })
        .catch(error => {
            console.log(error);
        });
}

async function getRiskyUserProfiles(token: string, users:RiskyUser[]): Promise<RiskyUserProfile[]> {
    let userinfos:RiskyUserProfile[] = [];
    for (let user of users) {
        let config: AxiosRequestConfig = {
            method: 'get',
            url: MS_GRAPH_ENDPOINT_BETA + 'users/' + user.userPrincipalName + '?$select=displayName,userPrincipalName,mobilePhone,otherMails',
            headers: {
              'Authorization': 'Bearer ' + token //the token is a variable which holds the token
            }
        }
        userinfos.push(await axios(config)
        .then(response => {
            // console.log(response.data);
            return response.data;
        })
        .catch(error => {
            console.log(error);
        }));
    }
    return userinfos;
}

async function resetPassword(token: string, user: RiskyUserProfile) {
    /* Get Password Profile ID */
    let config1: AxiosRequestConfig = {
        method: 'get',
        url: MS_GRAPH_ENDPOINT_BETA + 'users/' + user.userPrincipalName + '/authentication/passwordMethods',
        headers: {
          'Authorization': 'Bearer ' + token //the token is a variable which holds the token
        },
    }
    let passwordMethodId = await axios(config1)
        .then(response => {
            console.log(response.data);
            return response.data.value.id;
        })
        .catch(error => {
            console.log(error);
        });
    
    
    /* Reset Password */
    let config: AxiosRequestConfig = {
        method: 'post',
        url: MS_GRAPH_ENDPOINT_BETA + 'users/' + user.userPrincipalName + '/authentication/passwordMethods/' + passwordMethodId + '/resetPassword',
        headers: {
          'Authorization': 'Bearer ' + token //the token is a variable which holds the token
        },
        data: {
            newPassword: 'newPassword-value'
        }
    }
    await axios(config)
        .then(response => {
            console.log(response.data);
            return response.data;
        })
        .catch(error => {
            console.log(error);
        });
}

async function sendMail(token: string, userPrincipal: string, mailAddress: string) {
    let data = {
        "message": {
          "subject": "Risky User Alert",
          "body": {
            "contentType": "Text",
            "content": "IT WORKED"
          },
          "toRecipients": [
            {
              "emailAddress": {
                "address": mailAddress
              }
            }
          ]
        },
        "saveToSentItems": "false"
      }
    
    let config: AxiosRequestConfig = {
        method: 'post',
        url: MS_GRAPH_ENDPOINT + 'users/' + userPrincipal + '/sendMail',
        headers: {
          'Authorization': 'Bearer ' + token //the token is a variable which holds the token
        },
        data: data
    }
    console.log("MAIL", config);
    await axios(config)
        .then(response => {
            console.log(response.data);
            return response.data;
        })
        .catch(error => {
            console.log(error);
        });
}

async function sendSupportMail(token: string, user: RiskyUserProfile, riskyuser: RiskyUser) {
    let data = {
        "message": {
          "subject": "Risky User Alert",
          "body": {
            "contentType": "Text",
            "content": "Benutzer: " + user.userPrincipalName + " Datum: " + riskyuser.riskLastUpdatedDateTime + " Grund: " + riskyuser.riskDetail + "Alternative Kontaktmöglichkeiten: " + user.otherMails + ", " + user.mobilePhone
          },
          "toRecipients": [
            {
              "emailAddress": {
                "address": SUPPORT_EMAIL
              }
            }
          ]
        },
        "saveToSentItems": "false"
      }
    
    let config: AxiosRequestConfig = {
        method: 'post',
        url: MS_GRAPH_ENDPOINT + 'users/' + user.userPrincipalName + '/sendMail',
        headers: {
          'Authorization': 'Bearer ' + token //the token is a variable which holds the token
        },
        data: data
    }
    await axios(config)
        .then(response => {
            console.log(response.data);
            return response.data;
        })
        .catch(error => {
            console.log(error);
        });
}

async function sendSMS(mobilePhone: string, account: string) {
    twilio.messages
        .create({
            body: 'Risky User Alert! Für den Account: ' + account,
            from: TWILIO_NUMBER,
            to: mobilePhone
        })
        .then(message => console.log(message.sid));
}