import { NgModule } from '@angular/core';
import { RouterModule, Routes } from '@angular/router';
import { DashboardSixComponent } from './dashboard-six.component';



const routes: Routes = [{
  path: '',
  component: DashboardSixComponent,
}];

@NgModule({
  imports: [RouterModule.forChild(routes)],
  exports: [RouterModule]
})
export class DashboardSixRoutingModule { }
