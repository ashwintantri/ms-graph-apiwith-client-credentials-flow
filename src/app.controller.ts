import { Controller, Get } from '@nestjs/common';
import { AppService } from './app.service';
import { ClientSecretCredential } from "@azure/identity";
import { ConfidentialClientApplication } from '@azure/msal-node';
import { Client } from "@microsoft/microsoft-graph-client";
import { TokenCredentialAuthenticationProvider, TokenCredentialAuthenticationProviderOptions } from "@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials";
import "isomorphic-fetch";

@Controller()
export class AppController {
  constructor(private readonly appService: AppService) {}

  @Get()
  async getHello(): Promise<string> {
const tokenCredential = new ClientSecretCredential();
const options: TokenCredentialAuthenticationProviderOptions = { scopes: ["https://graph.microsoft.com/.default"] };

const authProvider = new TokenCredentialAuthenticationProvider(tokenCredential, options);
const client = Client.initWithMiddleware({
	debugLogging: true,
	authProvider: authProvider,
});

const event = {
  subject: 'Let\'s go for lunch',
  body: {
    contentType: 'HTML',
    content: 'Does next month work for you?'
  },
  start: {
      dateTime: '2019-03-10T12:00:00',
      timeZone: 'Pacific Standard Time'
  },
  end: {
      dateTime: '2019-03-10T14:00:00',
      timeZone: 'Pacific Standard Time'
  },
  location: {
      displayName: 'Harry\'s Bar'
  },
  attendees: [
    {
      emailAddress: {
        address: 'adelev@contoso.onmicrosoft.com',
        name: 'Adele Vance'
      },
      type: 'required'
    }
  ],
  isOnlineMeeting: true,
  onlineMeetingProvider: 'teamsForBusiness'
};
const res = await client.api("/users/f5fe735b-3941-44ff-800a-24b242d46dda/calendar/events").get();
    //const confidentialClientApplication = new ConfidentialClientApplication(clientConfig);
    //const accessToken = await getClientCredentialsToken(confidentialClientApplication)
    return res;
  }
}
