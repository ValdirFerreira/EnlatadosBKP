import { HttpClient } from '@angular/common/http';
import { Component, Input, OnInit } from '@angular/core';
import { ComunicacaoQuadroResumo } from 'src/app/models/ComunicacaoQuadroResumo/ComunicacaoQuadroResumo';
import { FiltroGlobalService } from 'src/app/services/filtro-global.service';

@Component({
    selector: 'grafico-propaganda',
    templateUrl: './grafico-propaganda.component.html',
    styleUrls: ['./grafico-propaganda.component.scss']
})
export class GraficoPropagandaComponent implements OnInit {

    @Input() codFoto: number = 0;

    @Input() graficoComunicacaoQuadroResumo: Array<ComunicacaoQuadroResumo>;

    constructor(
        public filtroService: FiltroGlobalService, private http: HttpClient
    ) { }

    ngOnInit(): void {
    }




    montaImagePropaganda(nFoto: any) {
        const mes = this.codFoto;
        const imageUrl = `assets/propaganda/${mes}_${nFoto}.png`;
        return imageUrl;
        // if (this.imageExists(imageUrl))
        //     return imageUrl;

        // else {

        //     const imageUrlNA = `assets/propaganda/NA.png`;
        //     return imageUrlNA
        // }

    }

    imageExists(url: string): Promise<boolean> {
        return this.http.head(url, { observe: 'response' })
            .toPromise()
            .then(response => response.status === 200)
            .catch(() => false);
    }


    itemSelected: string;
    openPopup(nFoto) {

        var mes = this.codFoto;
        this.itemSelected = "assets/propaganda/" + mes + "_" + nFoto + ".png";
        document.getElementById('popupWrapper').style.display = 'block';
    }

    closePopup() {
        document.getElementById('popupWrapper').style.display = 'none';
    }




}