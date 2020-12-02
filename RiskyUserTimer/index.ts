import { AzureFunction, Context, HttpRequest } from "@azure/functions";
import axios, { AxiosRequestConfig } from 'axios';
import qs = require('qs');

const SUPPORT_EMAIL = "jan.aebersold@be.ch"
const SUPPORT_SEND_MAIL = "jae@edubern-test.ch"
const APP_ID = "32ca135e-d663-4924-ab9c-dc4c33e3a819";
const APP_SECERET = "fWAVPrZ12Bg-4W~Y6_8Nx1m8Jlc18SO.GK";
const TENANT_ID = "a269cb9e-c4ae-4657-a96b-c288eafd30c2";

/*const SUPPORT_EMAIL = "thvome@gmail.com"
const SUPPORT_SEND_MAIL = "bob@projectxyzdemo.onmicrosoft.com"
const APP_ID = "4ecb1580-e2fc-4222-8900-3d34d994f3f5";
const APP_SECERET = "_0vlwIgPx3q1~8.e-Dq_SYLwj22snI6s4_";
const TENANT_ID = "2c315d01-a109-433f-b02c-583257b57e53";*/

const TOKEN_ENDPOINT = 'https://login.microsoftonline.com/' + TENANT_ID + '/oauth2/v2.0/token';
const MS_GRAPH_SCOPE = 'https://graph.microsoft.com/.default';
const MS_GRAPH_ENDPOINT = 'https://graph.microsoft.com/v1.0/';

const timerTrigger: AzureFunction = async function (context: Context, myTimer: any): Promise<void> {
    context.log('Timer trigger function processed a request.');

    // Set Default Header for Axios Requests
    axios.defaults.headers.post['Content-Type'] = 'application/x-www-form-urlencoded';

    // Get Token for MS Graph
    let token = await getToken();

    // Get detected Risks
    let riskdetections = await getRiskDetections(token);

    // Filter for only last 24h
    riskdetections = riskdetections.filter(isFrom2minutes)

    // Send Mail to defined User
    if (riskdetections.length > 0) {
        await sendSupportMail(token, riskdetections)
    }

    // Give back the detected Risks
    context.res = {
        // status: 200, /* Defaults to 200 */
        body: riskdetections
    };

};
export default timerTrigger;

function isFrom2minutes(riskd: RiskDetections) {
    console.log(riskd.id)
    let aDT = new Date(riskd.activityDateTime).getTime();
    let timeStamp = Math.round(new Date().getTime() / 1000);
    let timeStamp2minsago = timeStamp - (120);
    let is2minsago = aDT >= new Date(timeStamp2minsago*1000).getTime();
    console.log(is2minsago);
    if (is2minsago) {
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
        url: MS_GRAPH_ENDPOINT + 'identityProtection/riskDetections',
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
    let data = {
        "message": {
          "subject": "Risk Detections Alert",
          "body": {
            "contentType": "Text",
            "content": getEmailText(detections)
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
    let text: string = "Folgende Risk Detections wurden gefunden: \n\n";
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
        text += "EventType: " + detection.riskEventType + "\n";
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
    console.log(text);
    return text;
}