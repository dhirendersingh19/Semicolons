import { CommonModule } from '@angular/common';
import { NgModule, Component } from '@angular/core';
import { CommonMaterialModule } from '../common.material.module';
import { SlideViewComponent } from '../slide-view/slide-view.component';
import { SlidesTabComponent } from '../slides-tab/slides-tab.component';
import { StepperModule } from '../stepper/stepper.component';
import { TestCaseTabComponent } from '../test-case-tab/test-case-tab.component';
import { TestCaseViewComponent } from '../test-case-view/test-case-view.component';

@Component({
  selector: 'app-summary',
  templateUrl: './summary.component.html',
  styleUrls: ['./summary.component.scss']
})
export class SummaryComponent {

}


@NgModule({
  imports: [CommonMaterialModule, CommonModule, StepperModule],
  exports: [],
  declarations: [TestCaseTabComponent, TestCaseViewComponent, SummaryComponent]
})
export class SummaryModule {}
