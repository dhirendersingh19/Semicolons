import { ComponentFixture, TestBed } from '@angular/core/testing';

import { DashStepperComponent } from './dash-stepper.component';

describe('DashStepperComponent', () => {
  let component: DashStepperComponent;
  let fixture: ComponentFixture<DashStepperComponent>;

  beforeEach(async () => {
    await TestBed.configureTestingModule({
      declarations: [ DashStepperComponent ]
    })
    .compileComponents();

    fixture = TestBed.createComponent(DashStepperComponent);
    component = fixture.componentInstance;
    fixture.detectChanges();
  });

  it('should create', () => {
    expect(component).toBeTruthy();
  });
});
