import { Component, EventEmitter, Output } from '@angular/core';
import { MatSnackBar } from '@angular/material/snack-bar';
// import { MatSnackB}

@Component({
  selector: 'app-snackbar',
  templateUrl: './snackbar.component.html',
  styleUrls: ['./snackbar.component.scss']
})
export class SnackbarComponent {
  @Output() semit = new EventEmitter<any>();

  message: string = "";
  action: string = "";

  sendSnackbar(action: string) {
    this.semit.emit(this.openSnackBar);
  }

  constructor(private _snackBar: MatSnackBar) {}

  openSnackBar(message: string ="Analyzing Data", action: string = "Close") {
    this._snackBar.open(message, action);
  }

  // openSnackBar(message.value, action.value)
  // {

  // }
}
