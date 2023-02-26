import { STEPPER_GLOBAL_OPTIONS } from '@angular/cdk/stepper';
import { NgModule, Component, Input } from '@angular/core';
import { CommonMaterialModule } from '../common.material.module';

@Component({
  selector: 'app-stepper',
  templateUrl: './stepper.component.html',
  styleUrls: ['./stepper.component.scss']
})
export class StepperComponent {
  @Input() step = 1
  steps = [
    {
      completed: false,
      label: 'Upload transcript',
      state: 'upload'
    },
    {
      completed: false,
      label: 'Get slides',
      state: 'slides'
    },
    {
      completed: false,
      label: 'Get videos',
      state: 'videos'
    },
    {
      completed: false,
      label: 'Fill notes',
      state: 'notes'
    },
    {
      completed: false,
      label: 'Generate test cases',
      state: 'tests'
    }
  ]

  ngOnChanges() {
    console.log('step change', this.step, this.steps)
    const stepIndex = this.step - 1
    this.steps.map((value, index) => {
      if (index <= stepIndex) {
        value.completed = true
      } else {
        value.completed = false
      }
    })
  }
}

@NgModule({
  declarations: [StepperComponent],
  imports: [CommonMaterialModule],
  exports: [StepperComponent]
})
export class StepperModule {}