import { Component } from '@angular/core';

@Component({
  selector: 'app-google-auth',
  templateUrl: './google-auth.component.html',
  styleUrls: ['./google-auth.component.scss']
})
export class GoogleAuthComponent {
  // authInstance: any;
  // gapiSetup: boolean = false;
  // user: any;
  // error: any;
  // async initGoogleAuth(): Promise<void> {
  //   //  Create a new Promise where the resolve 
  //   // function is the callback passed to gapi.load
  //   const pload = new Promise((resolve) => {
  //     gapi.load('auth2', resolve);
  //   });

  //   // When the first promise resolves, it means we have gapi
  //   // loaded and that we can call gapi.init
  //   return pload.then(async () => {
  //     await gapi.auth2
  //       .init({ client_id: 'YOUR_GOOGLE_KEY' })
  //       .then(auth => {
  //         this.gapiSetup = true;
  //         this.authInstance = auth;
  //       });
  //   });
  // }

  // async authenticate(): Promise<gapi.auth2.GoogleUser> {
  //   // Initialize gapi if not done yet
  //   if (!this.gapiSetup) {
  //     await this.initGoogleAuth();
  //   }

  //   // Resolve or reject signin Promise
  //   return new Promise(async () => {
  //     await this.authInstance.signIn().then(
  //       (user: any) => this.user = user,
  //       (error: any) => this.error = error);
  //   });
  // }
}
