import { Injectable } from '@angular/core';
import { environment } from 'src/environments/environment';

@Injectable({
  providedIn: 'root'
})
export class MenuService {

  public activeMenu: number = 1;

  public menuSelecao: string = "1";

  public nomePage:string="";

  constructor() { }


}