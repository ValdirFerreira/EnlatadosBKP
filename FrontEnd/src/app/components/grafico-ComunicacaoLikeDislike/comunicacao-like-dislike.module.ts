import { NgModule } from '@angular/core';
import { CommonModule } from '@angular/common';
import { ComunicacaoLikeDislikeComponent } from './comunicacao-like-dislike.component';
import { NgSelectModule } from '@ng-select/ng-select';
import { SelectImageModule } from '../select-image/select-image.module';
import { FormsModule } from '@angular/forms';

@NgModule({
  declarations: [
    ComunicacaoLikeDislikeComponent
  ],
  imports: [
     CommonModule,
       NgSelectModule,
       SelectImageModule,
       FormsModule ,
  ],
  exports: [
    ComunicacaoLikeDislikeComponent
  ]
})

export class ComunicacaoLikeDislikeModule { }