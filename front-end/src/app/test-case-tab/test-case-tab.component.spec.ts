import { ComponentFixture, TestBed } from '@angular/core/testing';

import { TestCaseTabComponent } from './test-case-tab.component';

describe('TestCaseTabComponent', () => {
  let component: TestCaseTabComponent;
  let fixture: ComponentFixture<TestCaseTabComponent>;

  beforeEach(async () => {
    await TestBed.configureTestingModule({
      declarations: [ TestCaseTabComponent ]
    })
    .compileComponents();

    fixture = TestBed.createComponent(TestCaseTabComponent);
    component = fixture.componentInstance;
    fixture.detectChanges();
  });

  it('should create', () => {
    expect(component).toBeTruthy();
  });
});
