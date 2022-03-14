import { Injectable } from '@nestjs/common';
import * as msal from '@azure/msal-node';

@Injectable()
export class AppService {

  private msalClient: any;

  private config = {
    auth: {
      clientId: "855b2176-d34b-47ca-b8df-2d533a29888d",
      authority: "https://login.microsoftonline.com/49566b6c-6d19-4c4e-a0c5-cab837c0405c",
      clientSecret: "N4h7Q~IFJbG~06zKVzrVKY.DAJJHz93UcAG6U"
    },
    scopes: ["user.read","calendars.readwrite","mailboxsettings.read"],
    redirectUri: "http://localhost:4000"
  };


  constructor() {
    const msalConfig = {
      auth: this.config.auth,
      system: {
        loggerOptions: {
          loggerCallback(loglevel, message, containsPii) {
            console.log(message);
          },
          piiLoggingEnabled: false,
          logLevel: msal.LogLevel.Verbose,
        }
      }
    };

    // Create msal application object
    this.msalClient = new msal.ConfidentialClientApplication(msalConfig);
  }

  async getSignIn(res: any) {
    const urlParameters = {
      scopes: this.config.scopes,
      redirectUri: this.config.redirectUri
    };
    const authUrl = await this.msalClient.getAuthCodeUrl(urlParameters);
    res.redirect(authUrl);
  }

  async getCallback(res: any, req: any ){
    const tokenRequest = {
      code: req.query.code,
      scopes: this.config.scopes,
      redirectUri: this.config.redirectUri    
    }
    const response = await this.msalClient.acquireTokenByCode(tokenRequest);
    req.flash('error_msg', {
      message: 'Access token',
      debug: response.accessToken
    });
    res.redirect('/');
  }
}
