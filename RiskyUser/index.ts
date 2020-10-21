import { AzureFunction, Context, HttpRequest } from "@azure/functions"
import axios from 'axios';
import qs = require('qs');

const httpTrigger: AzureFunction = async function (context: Context, req: HttpRequest): Promise<void> {
    context.log('HTTP trigger function processed a request.');

    const APP_ID = "";
    const APP_SECERET = "";
    const TENANT_ID = "";
    const TOKEN_ENDPOINT ='https://login.microsoftonline.com/' + TENANT_ID + '/oauth2/v2.0/token';
    const MS_GRAPH_SCOPE = 'https://graph.microsoft.com/.default';
    

    
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
            return response.data;
        })
        .catch(error => {
            console.log(error);
        });

    const name = (req.query.name || (req.body && req.body.name));
    const responseMessage = name
        ? "Hello, " + name + ". This HTTP triggered function executed successfully."
        : "This HTTP triggered function executed successfully. Pass a name in the query string or in the request body for a personalized response.";

    context.res = {
        // status: 200, /* Defaults to 200 */
        body: token
    };

};

export default httpTrigger;