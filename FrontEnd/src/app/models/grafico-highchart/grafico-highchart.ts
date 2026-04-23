export class GraficoLinhasHighChartModel {
  type: string;
  name: string;
  data: Array<DataModel>;
  marker: any;
  tipoDados: number;

}

export class GraficoLinhasModel {
  Periodos: Array<string>;
  Grafico: Array<GraficoLinhasHighChartModel>
  Cores: Array<string>;
  Bases: Array<string>;
}

export class DataModel {
  name: string;
  y: number;
  periodo: string;
  valorbase: number;
  sig: string;
  tipoDados: number;
  media:number;
  baseminima:string;
}


