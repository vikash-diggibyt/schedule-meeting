import { Component, Inject, OnDestroy, OnInit } from '@angular/core';
import * as microsoftTeams from '@microsoft/teams-js';
import { MsalBroadcastService, MsalGuardConfiguration, MsalService, MSAL_GUARD_CONFIG } from '@azure/msal-angular';
import { InteractionStatus, RedirectRequest } from '@azure/msal-browser';
import { filter, Subject, takeUntil } from 'rxjs';
import { environment } from 'src/environments/environment';
import { AzureAdDemoService } from './azure-ad-demo.service';
import { Client } from '@microsoft/microsoft-graph-client';
import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';


@Component({
  selector: 'app-root',
  templateUrl: './app.component.html',
  styleUrls: ['./app.component.scss']
})
export class AppComponent implements OnInit,OnDestroy {
  isUserLoggedIn:boolean=false;
  userName?:string='';
  private readonly _destroy=new Subject<void>();
  constructor(@Inject(MSAL_GUARD_CONFIG) private msalGuardConfig:MsalGuardConfiguration,
  private msalBroadCastService:MsalBroadcastService,
  private authService:MsalService,private azureAdDemoSerice:AzureAdDemoService)
  {

  }
  ngOnInit(): void {
    this.msalBroadCastService.inProgress$.pipe
    (filter((interactionStatus:InteractionStatus)=>
    interactionStatus==InteractionStatus.None),
    takeUntil(this._destroy))
    .subscribe(x=>
      {
        this.isUserLoggedIn=this.authService.instance.getAllAccounts().length>0;
      
        if(this.isUserLoggedIn)
        {
          this.userName = this.authService.instance.getAllAccounts()[0].name;
        }
        this.azureAdDemoSerice.isUserLoggedIn.next(this.isUserLoggedIn);
      })
  }
  ngOnDestroy(): void {
   this._destroy.next(undefined);
   this._destroy.complete();
  }
  login()
  {
    if(this.msalGuardConfig.authRequest)
    {
      this.authService.loginRedirect({...this.msalGuardConfig.authRequest} as RedirectRequest)
    }
    else
    {
      this.authService.loginRedirect();
    }
  }
  logout()
  {
    this.authService.logoutRedirect({postLogoutRedirectUri:environment.postLogoutUrl});
  }


// // Authenticate and initialize the Graph client
// authenticateAndInitializeGraphClient() {
//   // Check if the user is authenticated
//   if (this.authService.getAccount()) {
//     // Get the access token
//     const accessToken = this.authService.acquireTokenSilent({
//       // Provide the appropriate scopes for Microsoft Graph API
//       scopes: ['user.read', 'calendars.readWrite']
//     }).then(response => {
//       // Initialize the Graph client with the access token
//       const client = Client.init({
//         authProvider: done => {
//           done(null, response.accessToken);
//         }
//       });

//       // Use the Graph client to create or join online meetings
//       this.createOrJoinOnlineMeeting(client);
//     }).catch(error => {
//       console.error('Failed to acquire access token:', error);
//     });
//   } else {
//     // Redirect the user to sign in if not authenticated
//     this.authService.loginRedirect();
//   }
// }

// // Create or join online meeting using the Graph client
// createOrJoinOnlineMeeting(client: Client) {
//   // Use the Graph client to create or join online meeting
//   // Refer to the Microsoft Graph API documentation for the appropriate API endpoint and request payload
//   // For example, to create a meeting, you can use the `POST /me/events` endpoint with the appropriate request payload
//   // For example:
//   const newMeeting: MicrosoftGraph.Event = {
//     subject: 'My Meeting',
//     start: {
//       dateTime: '2023-04-18T09:00:00',
//       timeZone: 'UTC'
//     },
//     end: {
//       dateTime: '2023-04-18T10:00:00',
//       timeZone: 'UTC'
//     },
//     location: {
//       displayName: 'Online Meeting'
//     },
//     onlineMeeting: {
//       // Provide the appropriate meeting settings
//       // Refer to the Microsoft Graph API documentation for available options
//     }
//   };
//   client.api('/me/events').post(newMeeting)
//     .then(response => {
//       console.log('Meeting created successfully:', response);
//     }).catch(error => {
//       console.error('Failed to create meeting:', error);
//     });
// }

}
