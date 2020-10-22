import { AzureFunction, Context, HttpRequest } from "@azure/functions"
import axios, { AxiosRequestConfig } from 'axios';
import qs = require('qs');


const SUPPORT_EMAIL = "thomas@balderdash.ch"

const APP_ID = "";
const APP_SECERET = "";
const TENANT_ID = "";
const TOKEN_ENDPOINT = 'https://login.microsoftonline.com/' + TENANT_ID + '/oauth2/v2.0/token';
const MS_GRAPH_SCOPE = 'https://graph.microsoft.com/.default';
const MS_GRAPH_ENDPOINT = 'https://graph.microsoft.com/v1.0/';
const MS_GRAPH_ENDPOINT_BETA = 'https://graph.microsoft.com/beta/';

const httpTrigger: AzureFunction = async function (context: Context, req: HttpRequest): Promise<void> {
    context.log('HTTP trigger function processed a request.');

    axios.defaults.headers.post['Content-Type'] = 'application/x-www-form-urlencoded';

    let token = await getToken();    
    let riskyusers = await getRiskyUsers(token);
    let userinfos = await getRiskyUserProfiles(token, riskyusers);
    
    
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
            console.log(response.data);
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
            console.log(response.data);
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
            console.log(response.data);
            return response.data;
        })
        .catch(error => {
            console.log(error);
        }));
    }
    return userinfos;
}

async function sendMail(from:string, to:string) {
    
}

async function sendSMS() {
    
}