
import {Component} from '@angular/core';
import {FormControl, Validators} from '@angular/forms';
import { Router } from '@angular/router';

import { HttpClient } from '@angular/common/http';

/** @title Form field with error messages */
@Component({
  selector: 'app-user-login',
  templateUrl: './user-login.component.html',
  styleUrls: ['./user-login.component.scss']
})
export class UserLoginComponent {
  hide = true;
  email = new FormControl('', [Validators.required]);

  constructor(private router: Router, private http: HttpClient) {

  }
  getErrorMessage() {
    if (this.email.hasError('required')) {
      return 'You must enter a value';
    }

    return this.email.hasError('email') ? 'Not a valid email' : '';
  }

  login(username: string= "user1", password: string= "pass123") {

    // pack data to send to backend
    const formData = new FormData();
    formData.append('username', username);
    formData.append('password', password);
    
    // code to check if user entered correct credentials.
    // send post request to backend to check if user exists
    this.http.post('http://backend-url', formData).subscribe(response => {
      console.log(response);

      // check from response object if successfully logged in
      // set the JWT tokem sent by backend if successfully logged in.
      // localStorage.setItem('demo_assist_token', "Value of JWT token here");
      // route to dashboard page if logged in. else route to login page again
      this.router.navigateByUrl('dashboard')

    }, error => {
      console.error('Error while logging in: ', error);
    });

  }

  //Demo code on how to unpack response object
  //Use above login() method instead when backend is ready
  login2(email:string, pass: string)
  {
  //   console.log("Inside login 2");
  //   this.http.get('http://localhost:8080/').subscribe(response => {
  //     console.log(response);
  //     if("status" in response) {
  //       console.log(response.status);
  //     }
  //   }, error => {
  //     console.error("Error: ", error);
  //   })

    this.http.post('http://localhost:5000/login', {email, pass}).subscribe(response => {
      console.log(response);
    }, error => {
      console.log(error);
    })
    
    if(email != 'semicolon@outlook.com') { alert('Invalid Credentials'); return;}

    localStorage.setItem('semicolon@outlook.com', pass);

    this.router.navigate(['dashboard']);

  }
}