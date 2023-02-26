import { CommonModule } from '@angular/common';
import { Component, NgModule } from '@angular/core';
import { CommonMaterialModule } from '../common.material.module';
import { NotesTabComponent } from '../notes-tab/notes-tab.component';
import { NotesViewComponent } from '../notes-view/notes-view.component';
import { SlideComponent } from '../slide/slide.component';
import { StepperModule } from '../stepper/stepper.component';

@Component({
  selector: 'app-notes',
  templateUrl: './notes.component.html',
  styleUrls: ['./notes.component.scss']
})
export class NotesComponent {

}

@NgModule({
  imports: [CommonMaterialModule, CommonModule, StepperModule],
  exports: [],
  declarations: [NotesViewComponent, NotesTabComponent, NotesComponent]
})
export class NotesModule {}
