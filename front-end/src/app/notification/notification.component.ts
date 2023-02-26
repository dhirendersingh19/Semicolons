import { Component, EventEmitter, Output } from '@angular/core';
import { HelperService } from '../helper.service';

@Component({
  selector: 'app-notification',
  templateUrl: './notification.component.html',
  styleUrls: ['./notification.component.scss']
})
export class NotificationComponent {
 events: {api: string, time: string, success: boolean}[] = [
  // {
  //   api: '',
  //   time: '',
  //   success: true 
  // }
 ]
 @Output() update = new EventEmitter<number>()
 eventsLength = 0

 constructor(private helper: HelperService) {

 }

 ngOnInit() {
  this.helper.event.subscribe((e) => {
    this.events.push(e)
    this.update.emit(this.events.length)
  })
 }
}
