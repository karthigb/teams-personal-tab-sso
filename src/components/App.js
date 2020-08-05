// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

import React from 'react';
import './App.css';
import { BrowserRouter as Router, Route } from "react-router-dom";

import Privacy from "./Privacy";
import TermsOfUse from "./TermsOfUse";
import Tab from "./Tab";
import ConsentPopup from "./ConsentPopup";
import ClosePopup from "./ClosePopup";

/**
 * The main app which handles the initialization and routing
 * of the app.
 */
class App extends React.Component {

  render() {

      // Set app routings that don't require microsoft Teams
      // SDK functionality.  Show an error if trying to access the
      // Home page.
      if (window.parent === window.self) {
        return (
          <Router>
            <Route exact path="/privacy" component={Privacy} />
            <Route exact path="/termsofuse" component={TermsOfUse} />
            <Route exact path="/tab" component={TeamsHostError} />
            <Route exact path="/auth-start" component={ConsentPopup} />
            <Route exact path="/auth-end" component={ClosePopup} />
          </Router>        
        );
      } else {
        // Display the app home page hosted in Teams
        return (
          <Router>
            <Route exact path="/tab" component={Tab}/>
          </Router>
        );
      }
    }
}

/**
 * This component displays an error message in the
 * case when a page is not being hosted within Teams.
 */
class TeamsHostError extends React.Component {
  render() {
    return (
      <div>
        <h3 className="Error">Teams client host not found.</h3>
      </div>
    );
  }
}

export default App;