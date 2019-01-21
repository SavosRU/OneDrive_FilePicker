import config from './Config';
import { UserAgentApplication } from 'msal';

import React from 'react';
//import ReactDOM from 'react-dom';
//import PropTypes from 'prop-types';
//import loadScript from 'load-script';
import { GraphFileBrowser } from '@microsoft/file-browser';

export default class OneDriveFilePicker extends React.Component {

  constructor(props) {
    super(props);

    this.userAgentApplication = new UserAgentApplication(config.appId, null, null);

    var user = this.userAgentApplication.getUser();

    this.state = {
      accessToken: null,
      isAuthenticated: (user !== null),
      user: {},
      error: null
    };

    if (user) {
      // Enhance user object with data from Graph
      this.getUserProfile();
    }

    console.log("constructor():", this);
    this.Login = this.Login.bind(this);
    this.getAuthenticationToken = this.getAuthenticationToken.bind(this);

  }

  async Login() {
    console.log("login():", this);

    try {
      await this.userAgentApplication.loginPopup(config.scopes);
      await this.getUserProfile();
    }
    catch(err) {
      //var errParts = err.split('|');
      // error: { message: errParts[1], debug: errParts[0] }
      // this.setState({
      //   isAuthenticated: false,
      //   user: {},
      //   error: { message: err }
      // });
      console.log("login(): \n\t", err);
    }
  }

  Logout = () => {
    this.userAgentApplication.logout();
  }

  async getUserProfile() {
    try {
      // Get the access token silently
      // If the cache contains a non-expired token, this function
      // will just return the cached token. Otherwise, it will
      // make a request to the Azure OAuth endpoint to get a token

      var accessToken = await this.userAgentApplication.acquireTokenSilent(config.scopes);

      if (accessToken) {
        // TEMPORARY: Display the token in the error flash
        this.setState({
          accessToken: accessToken,
          isAuthenticated: true,
          error: { message: "Access token:", debug: accessToken }
        });
      }

      console.log("getUserProfile(): \n\t", this);

    }
    catch(err) {
      var errParts = err.split('|');
      this.setState({
        isAuthenticated: false,
        user: {},
        error: { message: errParts[1], debug: errParts[0] }
      });
    }
  }



  render() {
    return (
      <div>
        { this.state.isAuthenticated !== true && (
          <div>
            <h3 id="WelcomeMessage">Начало АВТОРИЗАЦИИ</h3>
            <button id="SignIn" onClick={this.Login}>Sign In</button>
          </div>
        ) }

        <br/><br/>
        <pre id="json"></pre>

        { this.state.isAuthenticated !== false && (
          <div>
          <h3>beginning of the FileBrowser</h3>
          <GraphFileBrowser
            getAuthenticationToken={this.getAuthenticationToken}
            onSuccess={this.onSuccess}
          />
          <span>end of the FileBrowser</span>
          </div>
        ) }
      </div>
    );
  }

  getAuthenticationToken = () => {
    //return Promise.resolve('<access_token>');
    return Promise.resolve(this.state.accessToken);
  }

  onSuccess = (keys) => {
      // the key of the item in the File Browser's internal cache
      // [key: string]: {
      //   endpoint: string, // the endpoint url the item was fetched from
      //   driveId?: string, // the identifier of the drive that contains the item
      //   itemId?: string   // the identifier of the item
      // }
      console.log('onSuccess(): keys are ', keys);
      console.log('onSuccess(): keys[0][driveItem_203] are ', keys[0]['driveItem_203']);
      let file = keys[0]['driveItem_203'];

      console.log('thumbUrl = ', file[0] + "/drive/items/" + file[2] + "/thumbnails?select=c100x100_Crop");
      //console.log('downloadUrl = ', file[0] + "/drive/items/" + file[2] + "/content");
      //console.log('downloadUrl = ', file[0] + "/drive/items/" + file[2] + "?select=id,@microsoft.graph.downloadUrl");
      console.log('downloadUrl = ', file[0] + "/drives/" + file[1] + "/items/" + file[2] + "?select=id,@microsoft.graph.downloadUrl");
  }

}
