import { NgModule } from '@angular/core';
import { RouterModule, Routes } from '@angular/router';
import { DashboardEightComponent } from './dashboard-eight.component';

const routes: Routes = [{
  path: '',
  component: DashboardEightComponent,
}];

@NgModule({
  imports: [RouterModule.forChild(routes)],
  exports: [RouterModule]
})
export class DashboardEightRoutingModule { }
