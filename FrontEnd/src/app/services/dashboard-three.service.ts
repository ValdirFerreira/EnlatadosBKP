import { HttpBackend, HttpClient, HttpHeaders } from '@angular/common/http';
import { Injectable } from '@angular/core';
import { environment } from 'src/environments/environment';
import { FiltroPadrao, FiltroPadraoFullLoad } from '../models/Filtros/FiltroPadrao';
import { GraficoColunaModel } from '../models/grafico-coluna/grafico-coluna';
import { GraficoFunil } from '../models/grafico-funil/GraficoFunil';
import { AuthService } from './auth.service';

@Injectable({
  providedIn: 'root'
})
export class DashBoardThreeService {

  constructor(public httpClient: HttpClient,
    public httpClient2: HttpClient,
    private handler: HttpBackend,
    private authService: AuthService,) {
    this.httpClient2 = new HttpClient(handler);
  }

  private readonly baseUrl = environment["endPoint"];

  carregarGraficoComparativoExperiencia(filtro: FiltroPadrao) {
    return this.httpClient.post<GraficoFunil>(
      `${this.baseUrl}/DashBoardThree/CarregarGraficoComparativoExperiencia/`,
      filtro
    );
  }
 
  
}
