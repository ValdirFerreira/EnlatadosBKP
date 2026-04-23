import { HttpBackend, HttpClient, HttpHeaders } from '@angular/common/http';
import { Injectable } from '@angular/core';
import { environment } from 'src/environments/environment';
import { FiltroPadrao } from '../models/Filtros/FiltroPadrao';
import { GraficoLinhasModel } from '../models/grafico-highchart/grafico-highchart';
import { TabelaPadraoAdHoc } from '../models/tabela-padrao/TabelaPadraoAdHoc';
import { AuthService } from './auth.service';

@Injectable({
  providedIn: 'root'
})
export class DashBoardNineService {

  constructor(public httpClient: HttpClient,
    public httpClient2: HttpClient,
    private handler: HttpBackend,
    private authService: AuthService,) {
    this.httpClient2 = new HttpClient(handler);
  }

  private readonly baseUrl = environment["endPoint"];

  
 

  ImagemEvolutiva(filtro: FiltroPadrao) {
    return this.httpClient.post<GraficoLinhasModel>(
      `${this.baseUrl}/DashBoardNine/ImagemEvolutiva/`,
      filtro
    );
  }


  ImagemEvolutivaLinhas(filtro: FiltroPadrao) {
    return this.httpClient.post<GraficoLinhasModel>(
      `${this.baseUrl}/DashBoardNine/ImagemEvolutivaLinhas/`,
      filtro
    );
  }

  ImagemEvolutivaLinhas2(filtro: FiltroPadrao) {
    return this.httpClient.post<GraficoLinhasModel>(
      `${this.baseUrl}/DashBoardNine/ImagemEvolutivaLinhas2/`,
      filtro
    );
  }
  
  ImagemEvolutiTabelaAdHocAtributovaLinhas2(filtro: FiltroPadrao) {
    return this.httpClient.post<TabelaPadraoAdHoc>(
      `${this.baseUrl}/DashBoardNine/TabelaAdHocAtributo/`,
      filtro
    );
  }

  TabelaAdHocAtributoBloco6(filtro: FiltroPadrao) {
    return this.httpClient.post<TabelaPadraoAdHoc>(
      `${this.baseUrl}/DashBoardNine/TabelaAdHocAtributoBloco6/`,
      filtro
    );
  }


  TabelaAdHocAtributoBloco10(filtro: FiltroPadrao) {
    return this.httpClient.post<TabelaPadraoAdHoc>(
      `${this.baseUrl}/DashBoardNine/TabelaAdHocAtributoBloco10/`,
      filtro
    );
  }

  TabelaAdHocAtributoBloco2(filtro: FiltroPadrao) {
    return this.httpClient.post<TabelaPadraoAdHoc>(
      `${this.baseUrl}/DashBoardNine/TabelaAdHocAtributoBloco2/`,
      filtro
    );
  }
  



  TabelaAdHocAtributo(filtro: FiltroPadrao) {
    return this.httpClient.post<TabelaPadraoAdHoc>(
      `${this.baseUrl}/DashBoardNine/TabelaAdHocAtributo/`,
      filtro
    );
  }

  


  
 
  
}
