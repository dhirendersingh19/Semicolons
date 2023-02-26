import { Injectable } from '@angular/core';
import { BehaviorSubject, Observable } from 'rxjs';

@Injectable({
  providedIn: 'root'
})
export class TranferDataToViewService {
  private local_data = new BehaviorSubject(" ");

  constructor() { }

  getInfo(): Observable<string> {
    return this.local_data.asObservable();
  }

  setDataToShow(data: string)
  {
    this.local_data.next(data);
  }

  getDataToShow(): string
  {
    return this.local_data.getValue();
  }
}
