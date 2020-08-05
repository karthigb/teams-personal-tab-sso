// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

import React from 'react';

import crypto from 'crypto';
import * as microsoftTeams from "@microsoft/teams-js";

/**
 * This component is used to.... TODO
 */
class ConsentPopup extends React.Component {

    componentDidMount() {

      // Initialize the Microsoft Teams SDK
      microsoftTeams.initialize(window);

      // Get the user context from Teams and set it in the state
      microsoftTeams.getContext((context, error) => {

        let tenant = context['tid']; //Tenant ID of the logged in user
        let client_id = process.env.REACT_APP_AZURE_APP_REGISTRATION_ID; //Client ID of the Azure app registration ( may be from different tenant for multitenant apps)

        let queryParams = {
            tenant: `${tenant}`,
            client_id: `${client_id}`,
            response_type: "token",
            scope: "https://graph.microsoft.com/User.Read",
            redirect_uri: window.location.origin + "/auth-end",
            nonce: crypto.randomBytes(16).toString('base64')
        }
        
        let url = `https://login.microsoftonline.com/${tenant}/oauth2/v2.0/authorize?`;
        queryParams = new URLSearchParams(queryParams).toString();
        let authorizeEndpoint = url + queryParams;
        window.location.assign(authorizeEndpoint);

      });
    
    }    

    render() {
      return (
        <div>
          <h1>Redirecting to consent page...</h1>
        </div>
      );
    }
}

export default ConsentPopup;