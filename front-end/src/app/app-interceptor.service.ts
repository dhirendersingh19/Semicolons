import { HttpInterceptor, HttpRequest, HttpHandler, HttpEvent, HttpResponse } from '@angular/common/http';
import { Injectable } from '@angular/core';
import { tap, finalize, Observable } from 'rxjs';
import { HelperService } from './helper.service';

@Injectable()
export class AppInterceptorService implements HttpInterceptor {

  constructor(private helper: HelperService) { }

  intercept(req: HttpRequest<any>, next: HttpHandler): Observable<HttpEvent<any>> {
    this.helper.loader.emit({hide: false})
    let success = false
    let url: any = ''
    return next.handle(req).pipe(
      tap((event: HttpEvent<any>) => {
        console.log('http event', event)
        if (event instanceof HttpResponse) {
          url = event.url
          if (event.status >= 200 && event.status < 300) {
            success = true
          }
        }
      }),
      finalize(() => {
        this.helper.loader.emit({hide: true})
        
        // const now = new Date();
        // const year = now.getFullYear();
        // const month = (now.getMonth() + 1).toString().padStart(2, '0');
        // const day = now.getDate().toString().padStart(2, '0');
        // const hours = now.getHours().toString().padStart(2, '0');
        // const minutes = now.getMinutes().toString().padStart(2, '0');
        // const seconds = now.getSeconds().toString().padStart(2, '0');
        // const dateTime = `${year}-${month}-${day}-${hours}-${minutes}-${seconds}`;

        const now = new Date();
        const day = now.getDate();
        const suffix = this.getDaySuffix(day);
        const month = this.getMonthName(now.getMonth());
        const year = now.getFullYear().toString().slice(-2);
        const hours = now.getHours().toString().padStart(2, '0');
        const minutes = now.getMinutes().toString().padStart(2, '0');
        const dateTime = `${day}${suffix} ${month} ${year}, ${hours}:${minutes}`;

        this.helper.event.emit({
            success,
            api: url,
            time: dateTime
          })
      })
    )
  }
  
getDaySuffix(day: number) {
  if (day >= 11 && day <= 13) {
    return "th";
  }
  switch (day % 10) {
    case 1:
      return "st";
    case 2:
      return "nd";
    case 3:
      return "rd";
    default:
      return "th";
  }
}

getMonthName(month: number) {
  const monthNames = [
    "January", "February", "March", "April",
    "May", "June", "July", "August",
    "September", "October", "November", "December"
  ];
  return monthNames[month];
}

}
