import { NgModule } from '@angular/core';

import {MatGridListModule} from '@angular/material/grid-list';
import {MatListModule} from '@angular/material/list';

import {MatIconModule} from '@angular/material/icon';
import {MatButtonModule} from '@angular/material/button';

import {MatCardModule} from '@angular/material/card';
import {MatSnackBarModule} from '@angular/material/snack-bar';
import {MatToolbarModule} from '@angular/material/toolbar';
import {MatSidenavModule} from '@angular/material/sidenav'; 
import {MatDividerModule} from '@angular/material/divider';
import {MatBadgeModule} from '@angular/material/badge';
import {MatInputModule} from '@angular/material/input';
import { MatMenuModule } from '@angular/material/menu';

import {MatStepperModule} from '@angular/material/stepper'; 
import {MatProgressSpinnerModule} from '@angular/material/progress-spinner'; 
import {MatSelectModule} from '@angular/material/select';

@NgModule({
    imports: [
      MatToolbarModule,
      MatSidenavModule,
      MatDividerModule,
      MatBadgeModule,
      MatInputModule,
      MatMenuModule,
      MatGridListModule,
      MatListModule,
      MatIconModule,
      MatButtonModule,
      MatCardModule,
      MatProgressSpinnerModule,
      MatStepperModule,
      MatSelectModule,
      MatSnackBarModule
    ],
    exports: [
        MatToolbarModule,
        MatSidenavModule,
        MatDividerModule,
        MatBadgeModule,
        MatInputModule,
        MatMenuModule,
        MatGridListModule,
        MatListModule,
        MatIconModule,
        MatButtonModule,
        MatCardModule,
        MatProgressSpinnerModule,
        MatStepperModule,
        MatSelectModule
    ]
  })
  export class CommonMaterialModule { }