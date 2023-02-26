import { Component } from '@angular/core';
import { TranferDataToViewService } from '../shared/tranfer-data-to-view.service';

@Component({
  selector: 'app-notes-view',
  templateUrl: './notes-view.component.html',
  styleUrls: ['./notes-view.component.scss']
})
export class NotesViewComponent {
  transcript: string = "";

  constructor(private tranfer: TranferDataToViewService) {
    this.tranfer.getInfo()
    .subscribe(value => {
      this.transcript = value
      
    });
   }

  

  // ngOnChanges()
  // {
  //   this.transcript =  this.tranfer.getDataToShow();
  // }


}
