import { NgModule } from '@angular/core';
import { RouterModule, Routes } from '@angular/router';
// import { CommonMaterialModule } from './common.material.module';

import { LandingComponent } from './landing/landing.component';
import { NotesComponent } from './notes/notes.component';
import { SlideComponent } from './slide/slide.component';
import { SummaryComponent } from './summary/summary.component';
import { VideoComponent } from './video/video.component';
// import { FinalComponent } from './final/final.component';
import { NotFoundComponent } from './not-found/not-found.component';
import { UserLoginComponent } from './user-login/user-login.component';
import { UploadDataComponent } from './upload-data/upload-data.component';
import { AuthGuard } from './shared/auth.guard';
import { StepperComponent } from './stepper/stepper.component';
import { DashStepperComponent } from './dash-stepper/dash-stepper.component';

const routes: Routes = [
  {
  path: 'landing',
  component: LandingComponent,
  canActivate: [AuthGuard]
},
{
  path: 'dashboard',
  component: DashStepperComponent,
  canActivate: [AuthGuard]
},
{
  path: 'slides',
  component: SlideComponent,
  canActivate: [AuthGuard]
},
{
  path: 'videos',
  component: VideoComponent,
  canActivate: [AuthGuard]
},
{
  path: 'meet-notes',
  component: NotesComponent,
  canActivate: [AuthGuard]
},
{
  path: 'test-cases',
  component: SummaryComponent,
  canActivate: [AuthGuard]
},
// {
//   path: 'final',
//   component: FinalComponent
// },
{
  path: '',
  redirectTo: 'landing',
  pathMatch: 'full',
},
// {
//   path: 'steper',
//   component: DashStepperComponent,
//   canActivate: [AuthGuard]
// },
{
  path: 'login',
  component: UserLoginComponent,
  // outlet: 'global'
},
{
  path: '**',
  component: NotFoundComponent
}];

@NgModule({
  imports: [RouterModule.forRoot(routes)],
  exports: [RouterModule]
})
export class AppRoutingModule { }
