// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

import React from 'react';
import './App.css';
import * as microsoftTeams from "@microsoft/teams-js";

/**
 * This component is used to display...
 * about tab.
 */
class ClosePopup extends React.Component {

    notifyTab(){
      microsoftTeams.initialize(window);

      let hashParams = this.getHashParameters();

      if (hashParams["access_token"]){
        microsoftTeams.authentication.notifySuccess(hashParams["access_token"]);
      } else {
        microsoftTeams.authentication.notifyFailure("Consent failed");
      }
    }

    getHashParameters() {
      let hashParams = {};
      window.location.hash.substr(1).split("&").forEach(function(item) {
        let [key,value] = item.split('=');
        hashParams[key] = decodeURIComponent(value);
      });
      return hashParams;
  }    

    render() {
      this.notifyTab()
      return (
        <div>
          <h1>Consent flow complete.</h1>
        </div>
      );
    }
}

export default ClosePopup;