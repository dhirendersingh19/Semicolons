import { ComponentFixture, TestBed } from '@angular/core/testing';

import { TestCaseViewComponent } from './test-case-view.component';

describe('TestCaseViewComponent', () => {
  let component: TestCaseViewComponent;
  let fixture: ComponentFixture<TestCaseViewComponent>;

  beforeEach(async () => {
    await TestBed.configureTestingModule({
      declarations: [ TestCaseViewComponent ]
    })
    .compileComponents();

    fixture = TestBed.createComponent(TestCaseViewComponent);
    component = fixture.componentInstance;
    fixture.detectChanges();
  });

  it('should create', () => {
    expect(component).toBeTruthy();
  });
});
