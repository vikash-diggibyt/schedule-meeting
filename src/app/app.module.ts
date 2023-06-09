import { HTTP_INTERCEPTORS,HttpClientModule } from '@angular/common/http';
import { CUSTOM_ELEMENTS_SCHEMA, NgModule } from '@angular/core';
import { BrowserModule } from '@angular/platform-browser';
import { MsalGuard, MsalInterceptor, MsalModule, MsalRedirectComponent } from '@azure/msal-angular';
import { InteractionType, PublicClientApplication } from '@azure/msal-browser';
import { MatToolbarModule } from '@angular/material/toolbar';
import { MatButtonModule } from '@angular/material/button';
import {MatCardModule} from '@angular/material/card';
import {MatListModule} from '@angular/material/list';
import {MatDividerModule} from '@angular/material/divider';
import { AppRoutingModule } from './app-routing.module';
import { AppComponent } from './app.component';
import { ProfileComponent } from './profile/profile.component';
import { HomeComponent } from './home/home.component';
import { AzureAdDemoService } from './azure-ad-demo.service';
import { ReportComponent } from './report/report.component';
import {MatSnackBarModule} from '@angular/material/snack-bar';
import { BrowserAnimationsModule } from '@angular/platform-browser/animations';

const isIE=window.navigator.userAgent.indexOf('MSIE')>-1
||window.navigator.userAgent.indexOf('Trident/')>-1
@NgModule({
  declarations: [
    AppComponent,
    ProfileComponent,
    HomeComponent,
    ReportComponent
  ],
  imports: [
    BrowserModule,
    BrowserAnimationsModule,
    MatSnackBarModule,
    HttpClientModule,
    AppRoutingModule,
    MatToolbarModule,
    MatButtonModule,
    MatCardModule,
    MatListModule,
    MatDividerModule,
    MsalModule.forRoot(new PublicClientApplication
      (
        {
          auth:{
            clientId:'d7a7ca7b-b0c0-441c-ad6f-4b87cba5da55',
            redirectUri:'http://localhost:4200',
            authority:'https://login.microsoftonline.com/d074e917-c8d9-42cc-8986-2cd1bdfbd079'
          },
          cache:
          {
            cacheLocation:'localStorage',
            storeAuthStateInCookie:isIE
          }
        }
      ),
      {
        interactionType:InteractionType.Redirect,
        authRequest:{
          scopes:['user.read']
        }
      },
      {
        interactionType:InteractionType.Redirect,
        protectedResourceMap:new Map(
          [
            ['https://graph.microsoft.com/v1.0/me',['user.Read']],
            ['localhost',['api://apiUri/api.scope']]
          ]
        )
      }
      )
  ],
  providers: [{
    provide:HTTP_INTERCEPTORS,
    useClass:MsalInterceptor,
    multi:true
  },MsalGuard,AzureAdDemoService],
  bootstrap: [AppComponent,MsalRedirectComponent],
  schemas: [CUSTOM_ELEMENTS_SCHEMA]
})
export class AppModule { }
