import { NgModule } from '@angular/core';
import { BrowserModule } from '@angular/platform-browser';


import { AppRoutingModule } from './app-routing.module';
import { AppComponent } from './app.component';

// angular material
import { BrowserAnimationsModule } from '@angular/platform-browser/animations';

import { UserLoginComponent } from './user-login/user-login.component';

import {FormsModule, ReactiveFormsModule} from '@angular/forms';

import { GoogleAuthComponent } from './google-auth/google-auth.component';

// doc viewer
import { NgxDocViewerModule } from 'ngx-doc-viewer';

import { CommonMaterialModule } from './common.material.module';
// import { HttpClientModule } from '@angular/common/http';
import { TestCaseTabComponent } from './test-case-tab/test-case-tab.component';
import { TestCaseViewComponent } from './test-case-view/test-case-view.component';
// import { NotesViewComponent } from './notes-view/notes-view.component';
// import { NotesTabComponent } from './notes-tab/notes-tab.component';
// import { NotesComponent } from './notes/notes.component';

import { HttpClientModule, HTTP_INTERCEPTORS } from '@angular/common/http';
import { AppInterceptorService } from './app-interceptor.service';
import { STEPPER_GLOBAL_OPTIONS } from '@angular/cdk/stepper';
import { NotificationComponent } from './notification/notification.component';
import { DashStepperComponent } from './dash-stepper/dash-stepper.component';
import { VideoPlayerModule } from "./video-player/video-player.component";
import { CommonModule } from '@angular/common';
import { SnackbarComponent } from './snackbar/snackbar.component';

@NgModule({
    declarations: [
        AppComponent,
        UserLoginComponent,
        GoogleAuthComponent,
        NotificationComponent,
        DashStepperComponent,
        SnackbarComponent
    ],
    providers: [{
            provide: HTTP_INTERCEPTORS,
            useClass: AppInterceptorService,
            multi: true
        },
        // required for custom icons for the stepper
        {
            provide: STEPPER_GLOBAL_OPTIONS,
            useValue: { displayDefaultIndicatorType: false }
        }],
    bootstrap: [AppComponent],
    imports: [
        CommonModule,
        BrowserModule,
        HttpClientModule,
        AppRoutingModule,
        BrowserAnimationsModule,
        CommonMaterialModule,
        ReactiveFormsModule,
        FormsModule,
        NgxDocViewerModule,
        HttpClientModule,
        VideoPlayerModule,
    ]
})
export class AppModule { }
