import { Injectable } from '@angular/core';
import { Router } from '@angular/router';

@Injectable({
  providedIn: 'root'
})
export class AuthService {

  constructor(private router: Router) { }

  IsLoggedIn() 
  {
    if(!!localStorage.getItem('semicolon@outlook.com') && localStorage.getItem('semicolon@outlook.com') == '12345') return true;
    // return !!localStorage.getItem('demo_assist_token');
    return false;
  }

  // Call this method to perform logout throughout this application.
  logout()
  {
    // localStorage.removeItem('demo_assist_token');
    localStorage.removeItem('semicolon@outlook.com');
    this.router.navigate(['login']);
  }
}
