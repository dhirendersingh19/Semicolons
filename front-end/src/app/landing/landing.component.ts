import { NgModule } from '@angular/core';
import { CommonMaterialModule } from '../common.material.module';
import { ActivatedRoute } from '@angular/router';

import { Component } from '@angular/core';

@Component({
  selector: 'app-landing',
  templateUrl: './landing.component.html',
  styleUrls: ['./landing.component.scss']
})
export class LandingComponent {
  constructor(private route: ActivatedRoute) {
  }

  ngOnInit() {
    console.log("landing", this.route.snapshot.url.join('/'))
  }
}


@NgModule({
  imports: [
    CommonMaterialModule
  ],
  declarations: [LandingComponent],
})
export class LandingComponentModule { 
}
