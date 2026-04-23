import { Component } from '@angular/core';
import { NavigationEnd, Router } from '@angular/router';
import { TranslateService } from '@ngx-translate/core';
import { FiltroGlobalService } from './services/filtro-global.service';
import * as moment from 'moment';
import { AuthService } from './services/auth.service';
import { Session } from './pages/home/guards/session';

@Component({
  selector: 'app-root',
  templateUrl: './app.component.html',
  styleUrls: ['./app.component.scss']
})
export class AppComponent {
  title = 'abc';
  checkUserLogged: any;


  constructor(
    public translate: TranslateService,
    private router: Router,
    private filtroGlobalService: FiltroGlobalService,
    private authService: AuthService,
    private session: Session,
  ) {


    this.router.events.subscribe((event) => {
      translate.setDefaultLang('por');
      this.translate.use('por');
    });

  }

  ngOnInit() {


    var usuario = this.session.getSession();


    setTimeout(() => { this.callCheckUserLogged(); }, 200);
    this.checkUserLogged = setInterval(() => this.callCheckUserLogged(), 3000);

    this.router.events.subscribe((e) => {
      let ne = e as NavigationEnd;
      const arrNotUpdate = ['/', '/login'];

  
      if (ne != null && ne.url != null && !arrNotUpdate.includes(ne.url)) {
   
        this.callUpdateToken(false);
      }
    });
  }


  callCheckUserLogged() {

    if (this.router.url != '/' && !this.router.url.toLowerCase().includes('/home/login')) {
      let dtExpires = moment(this.filtroGlobalService.DtExpires);
      let dtNow = moment();
      let duration = dtExpires.diff(dtNow, 'seconds')

      if (duration <= 0) {
        this.sairPortal();
      }
      else if (duration <= 300) {
        let min = duration <= 60 ? 1 : Math.round(duration / 60);
      }
    }
  }


  callUpdateToken(blnShowMessage: boolean) {
    this.filtroGlobalService.DtExpires = moment().add(30, 'minutes').format();
  }



  sairPortal() {
    this.authService.limparDadosUsuario();
    this.router.navigate(['/login']);
  }


}


