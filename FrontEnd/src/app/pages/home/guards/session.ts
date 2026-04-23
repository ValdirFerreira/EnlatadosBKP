import { Injectable } from '@angular/core';
import { UsuarioModel } from 'src/app/models/usuario/UsuarioModel';


@Injectable()

export class Session {

    constructor() { }

     createSession(user: UsuarioModel) {
        localStorage.setItem('userNestleLeiteInfantil', JSON.stringify(user));
    }

    getSession() {
        let user = JSON.parse(localStorage.getItem('userNestleLeiteInfantil'));
        return user;
    }

    removeSession() {
        localStorage.removeItem('userNestleLeiteInfantil');
    }

    updateSession(user: UsuarioModel) {

        localStorage.removeItem('userNestleLeiteInfantil');

        localStorage.setItem('userNestleLeiteInfantil', JSON.stringify(user));

        let userSession = JSON.parse(localStorage.getItem('userNestleLeiteInfantil'));

        return userSession;
    }

    getCodUserSession() {
        let user = JSON.parse(localStorage.getItem('userNestleLeiteInfantil')) as UsuarioModel;
        return user.CodUser;
    }

    getUserSession() {
        let user = JSON.parse(localStorage.getItem('userNestleLeiteInfantil')) as UsuarioModel;
        return user
    }

    createSessionIPLogado(IP: string) {
        localStorage.removeItem('ipuserNestleLeiteInfantilLogado');
        localStorage.setItem('ipuserNestleLeiteInfantilLogado', IP);
    }

    getSessionIPLogado() {
        let IP = localStorage.getItem('ipuserNestleLeiteInfantilLogado');
        return IP;
    }



}