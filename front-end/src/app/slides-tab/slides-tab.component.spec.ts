import { ComponentFixture, TestBed } from '@angular/core/testing';

import { SlidesTabComponent } from './slides-tab.component';

describe('SlidesTabComponent', () => {
  let component: SlidesTabComponent;
  let fixture: ComponentFixture<SlidesTabComponent>;

  beforeEach(async () => {
    await TestBed.configureTestingModule({
      declarations: [ SlidesTabComponent ]
    })
    .compileComponents();

    fixture = TestBed.createComponent(SlidesTabComponent);
    component = fixture.componentInstance;
    fixture.detectChanges();
  });

  it('should create', () => {
    expect(component).toBeTruthy();
  });
});
