import { NgModule } from '@angular/core';
import { RouterModule, Routes } from '@angular/router';
import { DashboardNineComponent } from './dashboard-nine.component';


const routes: Routes = [{
  path: '',
  component: DashboardNineComponent,
}];

@NgModule({
  imports: [RouterModule.forChild(routes)],
  exports: [RouterModule]
})
export class DashboardNineRoutingModule { }
