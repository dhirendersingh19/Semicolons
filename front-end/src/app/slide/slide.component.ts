import { CommonModule } from '@angular/common';
import { NgModule, Component } from '@angular/core';
import { CommonMaterialModule } from '../common.material.module';
import { SlideViewComponent } from '../slide-view/slide-view.component';
import { SlidesTabComponent } from '../slides-tab/slides-tab.component';
import { StepperModule } from '../stepper/stepper.component';

@Component({
  selector: 'app-scan',
  templateUrl: './slide.component.html',
  styleUrls: ['./slide.component.scss']
})
export class SlideComponent {

}


@NgModule({
  imports: [CommonMaterialModule, CommonModule, StepperModule],
  exports: [],
  declarations: [SlidesTabComponent, SlideViewComponent, SlideComponent]
})
export class SlideModule {}