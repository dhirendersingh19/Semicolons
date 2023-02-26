import { NgModule, Component } from '@angular/core';
import { CommonMaterialModule } from '../common.material.module';
import { StepperModule } from '../stepper/stepper.component';

@Component({
  selector: 'app-upload-data',
  templateUrl: './upload-data.component.html',
  styleUrls: ['./upload-data.component.scss']
})
export class UploadDataComponent {
  onFileSelected(){
    
  }
}

@NgModule({
  declarations: [UploadDataComponent],
  imports: [CommonMaterialModule, StepperModule]
})
export class UploadModule {}