import { TestBed } from '@angular/core/testing';

import { TranferDataToViewService } from './tranfer-data-to-view.service';

describe('TranferDataToViewService', () => {
  let service: TranferDataToViewService;

  beforeEach(() => {
    TestBed.configureTestingModule({});
    service = TestBed.inject(TranferDataToViewService);
  });

  it('should be created', () => {
    expect(service).toBeTruthy();
  });
});
