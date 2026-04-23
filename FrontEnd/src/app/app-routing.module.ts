import { NgModule } from '@angular/core';
import { RouterModule, Routes } from '@angular/router';
import { AuthGuard } from './pages/home/guards/auth-guard.service';
import { LoginComponent } from './pages/home/login/login.component';
import { RecuperarSenhaComponent } from './pages/home/recuperar-senha/recuperar-senha.component';


const routes: Routes = [
  {
    path: '',
    pathMatch: 'full',
    redirectTo: 'home/login'
  },
  {
    path: 'recuperar-senha/:token',
    component: RecuperarSenhaComponent
  },

  { path: 'login', component: LoginComponent },
  { path: '', component: LoginComponent },
  {
    path: 'index',
    loadChildren: () => import('./pages/home/home.module').then(m => m.HomeModule),
  }
  , {
    path: 'home',
    loadChildren: () => import('./pages/menu-home/menu-home.module').then(m => m.MenuHomeModule),
    canActivate: [AuthGuard],
  }
  , {
    path: 'dashboard-one',
    loadChildren: () => import('./pages/dashboard-one/dashboard-one.module').then(m => m.DashboardOneModule),
    canActivate: [AuthGuard],
  }
  , {
    path: 'dashboard-awareness',
    loadChildren: () => import('./pages/dashboard-two/dashboard-two.module').then(m => m.DashboardTwoModule),
    canActivate: [AuthGuard],
  }
  , {
    path: 'dashboard-consideration',
    loadChildren: () => import('./pages/dashboard-eight/dashboard-eight.module').then(m => m.DashboardEightModule),
    canActivate: [AuthGuard],
  }
  , {
    path: 'dashboard-funil',
    loadChildren: () => import('./pages/dashboard-three/dashboard-three.module').then(m => m.DashboardThreeModule),
    canActivate: [AuthGuard],
  }
  , {
    path: 'dashboard-imagem-pura',
    loadChildren: () => import('./pages/dashboard-four/dashboard-four.module').then(m => m.DashboardFourModule),
    canActivate: [AuthGuard],
  }
  , {
    path: 'dashboard-posicionamento-marca',
    loadChildren: () => import('./pages/dashboard-five/dashboard-five.module').then(m => m.DashboardFiveModule),
    canActivate: [AuthGuard],
  }
  , {
    path: 'dashboard-brand-creator',
    loadChildren: () => import('./pages/dashboard-six/dashboard-six.module').then(m => m.DashboardSixModule),
    canActivate: [AuthGuard],
  }
  , {
    path: 'dashboard-comunicacao',
    loadChildren: () => import('./pages/dashboard-seven/dashboard-seven.module').then(m => m.DashboardSevenModule),
    canActivate: [AuthGuard],
  }
  , {
    path: 'dashboard-adhoc-section',
    loadChildren: () => import('./pages/dashboard-nine/dashboard-nine.module').then(m => m.DashboardNineModule),
    canActivate: [AuthGuard],
  },
  {
    path: 'usuario',
    loadChildren: () => import('./pages/usuario/usuario.module').then(m => m.UsuarioModule),
    canActivate: [AuthGuard],
  } 
];

@NgModule({
  imports: [RouterModule.forRoot(routes, { useHash: true })],
  exports: [RouterModule]
})
export class AppRoutingModule { }
