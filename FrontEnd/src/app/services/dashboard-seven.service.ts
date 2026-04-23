import { HttpBackend, HttpClient, HttpHeaders } from '@angular/common/http';
import { Injectable } from '@angular/core';
import { environment } from 'src/environments/environment';
import { ComunicacaoQuadroResumo } from '../models/ComunicacaoQuadroResumo/ComunicacaoQuadroResumo';
import { FiltroPadrao, FiltroPadraoFullLoad } from '../models/Filtros/FiltroPadrao';
import { GraficoColunaModel } from '../models/grafico-coluna/grafico-coluna';
import { GraficoImagemPosicionamento } from '../models/grafico-Imagem-posicionamento/GraficoImagemPosicionamento';
import { GraficoComunicacaoDiagnostico } from '../models/GraficoComunicacaoDiagnostico/GraficoComunicacaoDiagnostico';
import { GraficoComunicacaoRecall } from '../models/GraficoComunicacaoRecall/GraficoComunicacaoRecall';
import { GraficoComunicacaoVisto } from '../models/GraficoComunicacaoVisto/GraficoComunicacaoVisto';
import { AuthService } from './auth.service';

@Injectable({
  providedIn: 'root'
})
export class DashBoardSevenService {

  constructor(public httpClient: HttpClient,
    public httpClient2: HttpClient,
    private handler: HttpBackend,
    private authService: AuthService,) {
    this.httpClient2 = new HttpClient(handler);
  }

  private readonly baseUrl = environment["endPoint"];



  carregarGraficoComunicacaoVisto(filtro: FiltroPadrao) {
    return this.httpClient.post<Array<GraficoComunicacaoVisto>>(
      `${this.baseUrl}/DashBoardSeven/CarregarGraficoComunicacaoVisto/`,
      filtro
    );
  }

  carregarGraficoComunicacaoSource(filtro: FiltroPadrao) {
    return this.httpClient.post<Array<GraficoComunicacaoVisto>>(
      `${this.baseUrl}/DashBoardSeven/CarregarGraficoComunicacaoSource/`,
      filtro
    );
  }

  carregarGraficoComunicacaoRecall(filtro: FiltroPadrao) {
    return this.httpClient.post<Array<GraficoComunicacaoRecall>>(
      `${this.baseUrl}/DashBoardSeven/CarregarGraficoComunicacaoRecall/`,
      filtro
    );
  }
  
  carregarGraficoComunicacaoDiagnostico(filtro: FiltroPadrao) {
    return this.httpClient.post<Array<GraficoComunicacaoDiagnostico>>(
      `${this.baseUrl}/DashBoardSeven/CarregarGraficoComunicacaoDiagnostico/`,
      filtro
    );
  }

  carregarComunicacaoQuadroResumo(filtro: FiltroPadrao) {
    return this.httpClient.post<Array<ComunicacaoQuadroResumo>>(
      `${this.baseUrl}/DashBoardSeven/CarregarComunicacaoQuadroResumo/`,
      filtro
    );
  }

  
  
}
