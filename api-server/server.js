
const fetch = require('node-fetch');
const express = require('express');
const jwt_decode = require('jwt-decode');

const app = express();
const clientId = 'e780aa33-4b28-41ee-ba3c-0e4320ac955e';
const clientSecret = '_.4uTC237_2sOkzrVY1gygZ4JoXUF27-n8';
const graphScopes = 'https://graph.microsoft.com/User.Read';

let handleQueryError = function (err) {
    console.log("handleQueryError called: ", err);
    return new Response(JSON.stringify({
        code: 400,
        message: 'Stupid network Error'
    }));
};

app.get('/getGraphAccessToken', async (req,res) => {

    let tenantId = jwt_decode(req.query.ssoToken)['tid']; //Get the tenant ID from the decoded token
    let accessTokenEndpoint = `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`;

    //Create your access token query parameters
    //Learn more: https://docs.microsoft.com/en-us/azure/active-directory/develop/v2-oauth2-on-behalf-of-flow#first-case-access-token-request-with-a-shared-secret
    let accessTokenQueryParams = {
        grant_type: 'urn:ietf:params:oauth:grant-type:jwt-bearer',
        client_id: clientId,
        client_secret: clientSecret,
        assertion: req.query.ssoToken,
        scope: graphScopes,
        requested_token_use: "on_behalf_of",
    };

    accessTokenQueryParams = new URLSearchParams(accessTokenQueryParams).toString();

    let accessTokenReqOptions = {
        method:'POST',
        headers: {
            Accept: "application/json",
            "Content-Type": "application/x-www-form-urlencoded"},
        body: accessTokenQueryParams
    };

    let response = await fetch(accessTokenEndpoint,accessTokenReqOptions).catch(handleQueryError);

    let data = await response.json();
    console.log("Response data: ",data);
    if(!response.ok) {
        if( (data.error === 'invalid_grant') || (data.error === 'interaction_required') ) {
            //This is expected if it's the user's first time running the app ( user must consent ) or the admin requires MFA
            console.log("User must consent or perform MFA");
            res.status(403).json({ error: 'invalid_grant' }); //This error triggers the consent flow in the client.
        } else {
            //Unknown error
            console.log('Could not exchange access token for unknown reasons.');
            res.status(500).json({ error: 'Could not exchange access token' });
        }
    } else {
        //The on behalf of token exchange worked. Return the token (data object) to the client.
        res.send(data);
    }
});

// Handles any requests that don't match the ones above
app.get('*', (req,res) =>{
    console.log("Unhandled request: ",req);
    res.status(404).send("Path not defined");
});

const port = process.env.PORT || 5000;
app.listen(port);

console.log('API server is listening on port ' + port);