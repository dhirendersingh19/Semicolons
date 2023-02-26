import { EventEmitter, Injectable, Output } from '@angular/core';

@Injectable({
  providedIn: 'root'
})
export class HelperService {
  @Output() loader = new EventEmitter()
  @Output() event = new EventEmitter()

  constructor() { }
}
