import { AzureFunction, Context, HttpRequest } from "@azure/functions"
import axios, { AxiosRequestConfig } from 'axios';
import qs = require('qs');

const httpTrigger: AzureFunction = async function (context: Context, req: HttpRequest): Promise<void> {
    context.log('HTTP trigger function processed a request.');

    const APP_ID = "";
    const APP_SECERET = "";
    const TENANT_ID = "";
    const TOKEN_ENDPOINT = 'https://login.microsoftonline.com/' + TENANT_ID + '/oauth2/v2.0/token';
    const MS_GRAPH_SCOPE = 'https://graph.microsoft.com/.default';
    const MS_GRAPH_ENDPOINT = 'https://graph.microsoft.com/v1.0/';

    const postData = {
        client_id: APP_ID,
        scope: MS_GRAPH_SCOPE,
        client_secret: APP_SECERET,
        grant_type: 'client_credentials'
    };

    axios.defaults.headers.post['Content-Type'] = 'application/x-www-form-urlencoded';

    let token = await axios
        .post(TOKEN_ENDPOINT, qs.stringify(postData))
        .then(response => {
            console.log(response.data);
            return response.data.access_token;
        })
        .catch(error => {
            console.log(error);
        });
    
    let config: AxiosRequestConfig = {
        method: 'get',
        url: MS_GRAPH_ENDPOINT + 'identityProtection/riskyUsers',
        headers: {
          'Authorization': 'Bearer ' + token //the token is a variable which holds the token
        }
    }
    
    let riskyusers = await axios(config)
        .then(response => {
            console.log(response.data);
            return response.data.value;
        })
        .catch(error => {
            console.log(error);
        });
    
    context.res = {
        // status: 200, /* Defaults to 200 */
        body: riskyusers
    };

};

export default httpTrigger;