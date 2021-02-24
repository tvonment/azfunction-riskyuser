import { AzureFunction, Context, HttpRequest } from "@azure/functions"
import axios, { AxiosRequestConfig } from 'axios';
import qs = require('qs');

const SUPPORT_EMAILS = process.env["SUPPORT_EMAIL"].split(';');
const SUPPORT_SEND_MAIL = process.env["SUPPORT_SEND_MAIL"];
const APP_ID = process.env["APP_ID"];
const APP_SECERET = process.env["APP_SECRET"];
const TENANT_ID = process.env["TENANT_ID"];
const SCRIPT_LINK = process.env["SCRIPT_LINK"];

const TOKEN_ENDPOINT = 'https://login.microsoftonline.com/' + TENANT_ID + '/oauth2/v2.0/token';
const MS_GRAPH_SCOPE = 'https://graph.microsoft.com/.default';
const MS_GRAPH_ENDPOINT = 'https://graph.microsoft.com/v1.0/';

const httpTrigger: AzureFunction = async function (context: Context, req: HttpRequest): Promise<void> {
    context.log('HTTP trigger function processed a request.');

    // Set Default Header for Axios Requests
    axios.defaults.headers.post['Content-Type'] = 'application/x-www-form-urlencoded';

    // Get Token for MS Graph
    let token = await getToken();

    // Get detected Risks
    let riskdetections = await getRiskDetections(token);

    // Filter for only last 24h
    let is24riskdetections = riskdetections.filter(isFromLastDay)

    // Send Mail to defined User
    await sendSupportMail(token, is24riskdetections)

    // Give back the detected Risks
    context.res = {
        // status: 200, /* Defaults to 200 */
        body: riskdetections
    };

};
export default httpTrigger;

function isFromLastDay(riskd: RiskDetections) {
    console.log(riskd)
    let aDT = new Date(riskd.activityDateTime).getTime();
    let timeStamp = Math.round(new Date().getTime() / 1000);
    let timeStampYesterday = timeStamp - (24 * 3600);
    let is24 = aDT >= new Date(timeStampYesterday*1000).getTime();
    console.log(aDT);
    console.log(is24);
    if (is24) {
        return riskd;
    }
}

/**
 * Detected Risks
 */
class RiskDetections {
    id: string;
    riskType: boolean;
    riskEventType: false;
    riskLevel: string;
    riskState: string;
    riskDetail: string;
    activity: string;
    activityDateTime: string;
    userDisplayName: string;
    userPrincipalName: string;
    additionalInfo: string;
    location: {
        city: string;
        state: string;
        countryOrRegion: string;
    }
}

/**
 * Get Token for MS Graph
 */
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

/**
 * Get detected Risks
 * @param token Token to authenticate through MS Graph
 */
async function getRiskDetections(token:string): Promise<RiskDetections[]> {
    let config: AxiosRequestConfig = {
        method: 'get',
        url: MS_GRAPH_ENDPOINT + "identityProtection/riskDetections?$orderby=activityDateTime desc&$filter=riskLevel ne 'low'",
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

/**
 * Send Mail to defined MAIL
 * @param token Token to authenticate through MS Graph
 * @param detections Information about the detected Risks
 */
async function sendSupportMail(token: string, detections: RiskDetections[]) {
    let toRecipients = [];
    SUPPORT_EMAILS.forEach(mail => {
        toRecipients.push({
            "emailAddress": {
              "address": mail
            }
          })
    });
    
    let data = {
        "message": {
          "subject": "Risk Detections Alert",
          "body": {
            "contentType": "Text",
            "content": getEmailText(detections)
          },
          "toRecipients": toRecipients
        },
        "saveToSentItems": "false"
      }
    
    let config: AxiosRequestConfig = {
        method: 'post',
        url: MS_GRAPH_ENDPOINT + 'users/' + SUPPORT_SEND_MAIL + '/sendMail',
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

/**
 * Combine all the required Information for the Mail body
 * @param detections Information about the detected Risks
 */
function getEmailText(detections: RiskDetections[]): string {
    let text: string = "";
    text += "Risk Detections wurden gefunden: \n";
    text += "Script Time: " + new Date().toLocaleString() + " \n";
    text += "\n";
    detections.forEach(detection => {
        let date = new Date(detection.activityDateTime);
        text += "ID: " + detection.id + "\n";
        text += "Detection Zeit: " + 
            date.getDate() + "-" + 
            (date.getMonth() + 1) + "-" + 
            date.getFullYear() + " " + 
            date.getHours() + ":" + 
            date.getMinutes() + ":" + 
            date.getSeconds()  + "\n";
        text += "Risk Level: " + detection.riskLevel + "\n";
        text += "Risk Activity: " + detection.activity + "\n";
        text += "Risk Event Type: " + detection.riskEventType + "\n";
        text += "Risk Event Detail: " + detection.riskDetail + "\n";
        text += "User: " + detection.userDisplayName + "\n";
        if (detection.location) {
            if (detection.location.city) {
                text += "Location: " + detection.location.city + " ";
            }
            if (detection.location.state) {
                text += detection.location.state + " ";
            }
            if (detection.location.countryOrRegion) {
                text += detection.location.countryOrRegion; 
            }
            text += "\n";
        }
        if (detection.additionalInfo) {
            text += "Zusätzliche Informationen: \n"
            let addInfoList:[{Key:string, Value:string}] = JSON.parse(detection.additionalInfo);
            addInfoList.forEach(element => {
                text += " -> " + element.Key
                text += ": " + element.Value + "\n"
            });
        }
        text += "\n\n"
    });
    text += "Link to Detection Alerts: https://portal.azure.com/#blade/Microsoft_AAD_IAM/RiskyUsersBlade";
    text += "\n\n"
    text += "Scriptlink: " + SCRIPT_LINK;
    return text;
}