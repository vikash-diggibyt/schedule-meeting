import { Component, OnInit } from '@angular/core';
import { AzureAdDemoService } from '../azure-ad-demo.service';

@Component({
  selector: 'app-home',
  templateUrl: './home.component.html',
  styleUrls: ['./home.component.scss']
})
export class HomeComponent implements OnInit {
  isUserLoggedIn:boolean=false;
  constructor(private azureAdDemoService:AzureAdDemoService) { }

  ngOnInit(): void {
    this.azureAdDemoService.isUserLoggedIn.subscribe(
      x=>{
        this.isUserLoggedIn=x;
      }
    )
  }

  token(){
    const data = {

      subject: 'Plan summer company picnic',
      body: {
        contentType: 'HTML',
        content: 'Let\'s kick-start this event planning!'
      },
      start: {
        "dateTime": "2023-06-09T10:00:00",
          timeZone: 'Pacific Standard Time'
      },
      end: {
        "dateTime": "2023-06-09T12:00:00",
          timeZone: 'Pacific Standard Time'
      },
      attendees: [
        {
          emailAddress: {
            "address": "kabin.m@diggibyte.com",
            name: 'Dana Swope'
          },
          type: 'Required'
        },
        {
          emailAddress: {
            "address": "vikash.kumar@diggibyte.com",
            name: 'Alex Wilber'
          },
          type: 'Required'
        }
      ],
      location: {
        displayName: 'Conf Room 3; Fourth Coffee; Home Office',
        locationType: 'Default'
      },
      locations: [
        {
          displayName: 'Conf Room 3'
        },
        {
          displayName: 'Fourth Coffee',
          address: {
            street: '4567 Main St',
            city: 'Redmond',
            state: 'WA',
            countryOrRegion: 'US',
            postalCode: '32008'
          },
          coordinates: {
            latitude: 47.672,
            longitude: -102.103
          }
        },
        {
          displayName: 'Home Office'
        }
      ],
      allowNewTimeProposals: true
    };

    this.azureAdDemoService.schedulMeeting(data).subscribe((res:any)=>{ 
      console.log(res);
    })
  }

}
