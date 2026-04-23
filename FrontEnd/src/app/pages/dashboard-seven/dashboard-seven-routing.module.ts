import { NgModule } from '@angular/core';
import { RouterModule, Routes } from '@angular/router';
import { DashboardSevenComponent } from './dashboard-seven.component';

const routes: Routes = [{
  path: '',
  component: DashboardSevenComponent,
}];

@NgModule({
  imports: [RouterModule.forChild(routes)],
  exports: [RouterModule]
})
export class DashboardSevenRoutingModule { }
