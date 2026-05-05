import { HttpClient } from '@angular/common/http';
import { Component, Input, OnInit, OnChanges, SimpleChanges } from '@angular/core';
import { ComunicacaoQuadroResumo } from 'src/app/models/ComunicacaoQuadroResumo/ComunicacaoQuadroResumo';
import { FiltroGlobalService } from 'src/app/services/filtro-global.service';

@Component({
    selector: 'grafico-propaganda',
    templateUrl: './grafico-propaganda.component.html',
    styleUrls: ['./grafico-propaganda.component.scss']
})
export class GraficoPropagandaComponent implements OnInit, OnChanges {

    @Input() codFoto: number = 0;

    @Input() graficoComunicacaoQuadroResumo: Array<ComunicacaoQuadroResumo>;

    fotos: number[] = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12];

    imagensValidas: { [key: number]: boolean } = {};

    itemSelected: string;

    constructor(
        public filtroService: FiltroGlobalService,
        private http: HttpClient
    ) { }

    ngOnInit(): void {
        this.resetImagensValidas();
    }

    ngOnChanges(changes: SimpleChanges): void {
        // Quando codFoto mudar, reseta o controle para tentar carregar as novas imagens
        if (changes['codFoto']) {
            this.resetImagensValidas();
        }
    }

    resetImagensValidas() {
        this.imagensValidas = {};
        this.fotos.forEach(n => this.imagensValidas[n] = true);
    }

    montaImagePropaganda(nFoto: number): string {
        const mes = this.codFoto;
        return `assets/propaganda/${mes}_${nFoto}.png`;
    }

    onImageError(nFoto: number) {
        this.imagensValidas[nFoto] = false;
    }

    openPopup(nFoto: number) {
        const mes = this.codFoto;
        this.itemSelected = `assets/propaganda/${mes}_${nFoto}.png`;
        document.getElementById('popupWrapper').style.display = 'block';
    }

    closePopup() {
        document.getElementById('popupWrapper').style.display = 'none';
    }
}