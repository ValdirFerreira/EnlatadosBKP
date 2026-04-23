using Dapper;
using Entities.Filtros;
using Entities.Parametros;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Business.TrataDados
{
    public class TrataFiltros
    {

        public DynamicParameters MontaParametrosFiltroPadraoComunicacao(FiltroPadrao filtro)
        {
            var parametros = new DynamicParameters();
            //parametros.Add("@ParamListaTarget", ""); //parametros.Add("@ParamListaTarget", filtro.Target == null ? "" : string.Join(",", filtro.Target.Select(s => string.Concat(s.IdItem))));
            parametros.Add("@ParamListaTarget", "");
            //parametros.Add("@ParamListaTarget", (filtro.Target == null || filtro.Target.FirstOrDefault() == null)  ? "" : string.Join(",", filtro.Target.Select(s => string.Concat(s.IdItem))));
            parametros.Add("@ParamListaRegiao", filtro.Regiao == null ? "" : string.Join(",", filtro.Regiao.Select(s => string.Concat(s.IdItem))));
            parametros.Add("@ParamListaRegiaoTipo", filtro.Regiao == null ? "" : string.Join(",", filtro.Regiao.Select(s => string.Concat(s.Tipo))));
            parametros.Add("@ParamListaDemografico", filtro.Demografico == null ? "" : string.Join(",", filtro.Demografico.Select(s => string.Concat(s.IdItem))));
            parametros.Add("@ParamListaDemograficoTipo", filtro.Demografico == null ? "" : string.Join(",", filtro.Demografico.Select(s => string.Concat(s.Tipo))));
            parametros.Add("@ParamListaOnda", filtro.Onda == null ? "" : string.Join(",", filtro.Onda.Select(s => string.Concat(s.IdItem))));
            parametros.Add("@ParamListaMarca", filtro.Marca == null ? "" : string.Join(",", filtro.Marca.Select(s => string.Concat(s.IdItem))));

            parametros.Add("@ParamCodUser", filtro.CodUser);
            parametros.Add("@ParamCodIdioma", filtro.CodIdioma);
            parametros.Add("@ParamSequencia", filtro.Sequencia);
            //parametros.Add("@ParamBIA", filtro.ParamBIA);           

            return parametros;
        }


        public DynamicParameters MontaParametrosFiltroPadrao(FiltroPadrao filtro)
        {

            if (filtro.Onda[0] == null)
            {
                filtro.Onda = null;
            }

            var parametros = new DynamicParameters();
            //parametros.Add("@ParamListaTarget", ""); //parametros.Add("@ParamListaTarget", filtro.Target == null ? "" : string.Join(",", filtro.Target.Select(s => string.Concat(s.IdItem))));
            parametros.Add("@ParamListaTarget", filtro.Target[0] == null ? "" : string.Join(",", filtro.Target.Select(s => string.Concat(s.IdItem))));
            //parametros.Add("@ParamListaTarget", (filtro.Target == null || filtro.Target.FirstOrDefault() == null)  ? "" : string.Join(",", filtro.Target.Select(s => string.Concat(s.IdItem))));
            parametros.Add("@ParamListaRegiao", filtro.Regiao == null ? "" : string.Join(",", filtro.Regiao.Select(s => string.Concat(s.IdItem))));
            parametros.Add("@ParamListaRegiaoTipo", filtro.Regiao == null ? "" : string.Join(",", filtro.Regiao.Select(s => string.Concat(s.Tipo))));
            parametros.Add("@ParamListaDemografico", filtro.Demografico == null ? "" : string.Join(",", filtro.Demografico.Select(s => string.Concat(s.IdItem))));
            parametros.Add("@ParamListaDemograficoTipo", filtro.Demografico == null ? "" : string.Join(",", filtro.Demografico.Select(s => string.Concat(s.Tipo))));
            parametros.Add("@ParamListaOnda", filtro.Onda == null ? "" : string.Join(",", filtro.Onda.Select(s => string.Concat(s.IdItem))));
            parametros.Add("@ParamListaMarca", filtro.Marca == null ? "" : string.Join(",", filtro.Marca.Select(s => string.Concat(s.IdItem))));

            parametros.Add("@ParamCodUser", filtro.CodUser);
            parametros.Add("@ParamCodIdioma", filtro.CodIdioma);
            parametros.Add("@ParamSequencia", filtro.Sequencia);
            //parametros.Add("@ParamBIA", filtro.ParamBIA);           

            return parametros;
        }

        public DynamicParameters MontaParametrosFiltroPadraoDenominator(FiltroPadrao filtro)
        {

            if (filtro.Onda[0] == null)
            {
                filtro.Onda = null;
            }

            var parametros = new DynamicParameters();
            //parametros.Add("@ParamListaTarget", ""); //parametros.Add("@ParamListaTarget", filtro.Target == null ? "" : string.Join(",", filtro.Target.Select(s => string.Concat(s.IdItem))));
            parametros.Add("@ParamListaTarget", filtro.Target[0] == null ? "" : string.Join(",", filtro.Target.Select(s => string.Concat(s.IdItem))));
            parametros.Add("@ParamListaRegiao", filtro.Regiao == null ? "" : string.Join(",", filtro.Regiao.Select(s => string.Concat(s.IdItem))));
            parametros.Add("@ParamListaRegiaoTipo", filtro.Regiao == null ? "" : string.Join(",", filtro.Regiao.Select(s => string.Concat(s.Tipo))));
            parametros.Add("@ParamListaDemografico", filtro.Demografico == null ? "" : string.Join(",", filtro.Demografico.Select(s => string.Concat(s.IdItem))));
            parametros.Add("@ParamListaDemograficoTipo", filtro.Demografico == null ? "" : string.Join(",", filtro.Demografico.Select(s => string.Concat(s.Tipo))));
            parametros.Add("@ParamListaOnda", filtro.Onda == null ? "" : string.Join(",", filtro.Onda.Select(s => string.Concat(s.IdItem))));
            parametros.Add("@ParamListaMarca", filtro.Marca == null ? "" : string.Join(",", filtro.Marca.Select(s => string.Concat(s.IdItem))));

            parametros.Add("@ParamCodUser", filtro.CodUser);
            parametros.Add("@ParamCodIdioma", filtro.CodIdioma);
            parametros.Add("@ParamSequencia", filtro.Sequencia);
            //parametros.Add("@ParamDenominators", filtro.ParamDenominators);
            //parametros.Add("@ParamBIA", filtro.ParamBIA);           

            return parametros;
        }


        public DynamicParameters MontaParametrosFiltroPadraoImagem(FiltroPadrao filtro)
        {
            var parametros = new DynamicParameters();
            parametros.Add("@ParamListaTarget", filtro.Target[0] == null ? "" : string.Join(",", filtro.Target.Select(s => string.Concat(s.IdItem))));
            //parametros.Add("@ParamListaTarget", (filtro.Target == null || filtro.Target.FirstOrDefault() == null)  ? "" : string.Join(",", filtro.Target.Select(s => string.Concat(s.IdItem))));
            parametros.Add("@ParamListaRegiao", filtro.Regiao == null ? "" : string.Join(",", filtro.Regiao.Select(s => string.Concat(s.IdItem))));
            parametros.Add("@ParamListaRegiaoTipo", filtro.Regiao == null ? "" : string.Join(",", filtro.Regiao.Select(s => string.Concat(s.Tipo))));
            parametros.Add("@ParamListaDemografico", filtro.Demografico == null ? "" : string.Join(",", filtro.Demografico.Select(s => string.Concat(s.IdItem))));
            parametros.Add("@ParamListaDemograficoTipo", filtro.Demografico == null ? "" : string.Join(",", filtro.Demografico.Select(s => string.Concat(s.Tipo))));
            parametros.Add("@ParamListaOnda", filtro.Onda == null ? "" : string.Join(",", filtro.Onda.Select(s => string.Concat(s.IdItem))));
            parametros.Add("@ParamListaMarca", filtro.Marca == null ? "" : string.Join(",", filtro.Marca.Select(s => string.Concat(s.IdItem))));

            parametros.Add("@ParamCodUser", filtro.CodUser);
            parametros.Add("@ParamCodIdioma", filtro.CodIdioma);
            parametros.Add("@ParamSequencia", filtro.Sequencia);
            parametros.Add("@ParamBIA", filtro.ParamBIA);
            return parametros;
        }

        public DynamicParameters MontaParametrosFiltroPadraoImagemAdHoc(FiltroPadrao filtro)
        {
            var parametros = new DynamicParameters();
            parametros.Add("@ParamListaTarget", ""); //parametros.Add("@ParamListaTarget", filtro.Target == null ? "" : string.Join(",", filtro.Target.Select(s => string.Concat(s.IdItem))));
            //parametros.Add("@ParamListaTarget", (filtro.Target == null || filtro.Target.FirstOrDefault() == null)  ? "" : string.Join(",", filtro.Target.Select(s => string.Concat(s.IdItem))));
            parametros.Add("@ParamListaRegiao", filtro.Regiao == null ? "" : string.Join(",", filtro.Regiao.Select(s => string.Concat(s.IdItem))));
            parametros.Add("@ParamListaRegiaoTipo", filtro.Regiao == null ? "" : string.Join(",", filtro.Regiao.Select(s => string.Concat(s.Tipo))));
            parametros.Add("@ParamListaDemografico", filtro.Demografico == null ? "" : string.Join(",", filtro.Demografico.Select(s => string.Concat(s.IdItem))));
            parametros.Add("@ParamListaDemograficoTipo", filtro.Demografico == null ? "" : string.Join(",", filtro.Demografico.Select(s => string.Concat(s.Tipo))));
            parametros.Add("@ParamListaOnda", filtro.Onda == null ? "" : string.Join(",", filtro.Onda.Select(s => string.Concat(s.IdItem))));
            parametros.Add("@ParamListaMarca", filtro.Marca == null ? "" : string.Join(",", filtro.Marca.Select(s => string.Concat(s.IdItem))));

            parametros.Add("@ParamCodUser", filtro.CodUser);
            parametros.Add("@ParamCodIdioma", filtro.CodIdioma);
            parametros.Add("@ParamSequencia", filtro.Sequencia);
            parametros.Add("@ParamCodBloco", filtro.Atributo);
            return parametros;
        }

        public DynamicParameters MontaParametrosFiltroPadraoImagemAdHocCustom(FiltroPadrao filtro)
        {
            var parametros = new DynamicParameters();
            parametros.Add("@ParamListaTarget", ""); //parametros.Add("@ParamListaTarget", filtro.Target == null ? "" : string.Join(",", filtro.Target.Select(s => string.Concat(s.IdItem))));
            //parametros.Add("@ParamListaTarget", (filtro.Target == null || filtro.Target.FirstOrDefault() == null)  ? "" : string.Join(",", filtro.Target.Select(s => string.Concat(s.IdItem))));
            parametros.Add("@ParamListaRegiao", filtro.Regiao == null ? "" : string.Join(",", filtro.Regiao.Select(s => string.Concat(s.IdItem))));
            parametros.Add("@ParamListaRegiaoTipo", filtro.Regiao == null ? "" : string.Join(",", filtro.Regiao.Select(s => string.Concat(s.Tipo))));
            parametros.Add("@ParamListaDemografico", filtro.Demografico == null ? "" : string.Join(",", filtro.Demografico.Select(s => string.Concat(s.IdItem))));
            parametros.Add("@ParamListaDemograficoTipo", filtro.Demografico == null ? "" : string.Join(",", filtro.Demografico.Select(s => string.Concat(s.Tipo))));
            parametros.Add("@ParamListaOnda", filtro.Onda == null ? "" : string.Join(",", filtro.Onda.Select(s => string.Concat(s.IdItem))));
            parametros.Add("@ParamListaMarca", filtro.Marca == null ? "" : string.Join(",", filtro.Marca.Select(s => string.Concat(s.IdItem))));

            parametros.Add("@ParamCodUser", filtro.CodUser);
            parametros.Add("@ParamCodIdioma", filtro.CodIdioma);
            parametros.Add("@ParamSequencia", filtro.Sequencia);
            parametros.Add("@ParamBloco", filtro.Atributo);
            return parametros;
        }

        public DynamicParameters MontaParametrosFiltroPadraoImagemAdHocSemMarca(FiltroPadrao filtro)
        {
            var parametros = new DynamicParameters();
            parametros.Add("@ParamListaTarget", ""); //parametros.Add("@ParamListaTarget", filtro.Target == null ? "" : string.Join(",", filtro.Target.Select(s => string.Concat(s.IdItem))));
            //parametros.Add("@ParamListaTarget", (filtro.Target == null || filtro.Target.FirstOrDefault() == null)  ? "" : string.Join(",", filtro.Target.Select(s => string.Concat(s.IdItem))));
            parametros.Add("@ParamListaRegiao", filtro.Regiao == null ? "" : string.Join(",", filtro.Regiao.Select(s => string.Concat(s.IdItem))));
            parametros.Add("@ParamListaRegiaoTipo", filtro.Regiao == null ? "" : string.Join(",", filtro.Regiao.Select(s => string.Concat(s.Tipo))));
            parametros.Add("@ParamListaDemografico", filtro.Demografico == null ? "" : string.Join(",", filtro.Demografico.Select(s => string.Concat(s.IdItem))));
            parametros.Add("@ParamListaDemograficoTipo", filtro.Demografico == null ? "" : string.Join(",", filtro.Demografico.Select(s => string.Concat(s.Tipo))));
            parametros.Add("@ParamListaOnda", filtro.Onda == null ? "" : string.Join(",", filtro.Onda.Select(s => string.Concat(s.IdItem))));
            parametros.Add("@ParamListaMarca", "");

            parametros.Add("@ParamCodUser", filtro.CodUser);
            parametros.Add("@ParamCodIdioma", filtro.CodIdioma);
            parametros.Add("@ParamSequencia", filtro.Sequencia);
            parametros.Add("@ParamCodBloco", filtro.Atributo);
            return parametros;
        }



        public DynamicParameters MontaParametrosFiltroPadraoEEfeitos(FiltroPadrao filtro)
        {
            var parametros = new DynamicParameters();
            parametros.Add("@ParamListaTarget", ""); //parametros.Add("@ParamListaTarget", filtro.Target == null ? "" : string.Join(",", filtro.Target.Select(s => string.Concat(s.IdItem))));
            //parametros.Add("@ParamListaTarget", (filtro.Target == null || filtro.Target.FirstOrDefault() == null) ? "" : string.Join(",", filtro.Target.Select(s => string.Concat(s.IdItem))));
            parametros.Add("@ParamListaRegiao", filtro.Regiao == null ? "" : string.Join(",", filtro.Regiao.Select(s => string.Concat(s.IdItem))));
            parametros.Add("@ParamListaRegiaoTipo", filtro.Regiao == null ? "" : string.Join(",", filtro.Regiao.Select(s => string.Concat(s.Tipo))));
            parametros.Add("@ParamListaDemografico", filtro.Demografico == null ? "" : string.Join(",", filtro.Demografico.Select(s => string.Concat(s.IdItem))));
            parametros.Add("@ParamListaDemograficoTipo", filtro.Demografico == null ? "" : string.Join(",", filtro.Demografico.Select(s => string.Concat(s.Tipo))));
            parametros.Add("@ParamListaOnda", filtro.Onda == null ? "" : string.Join(",", filtro.Onda.Select(s => string.Concat(s.IdItem))));
            parametros.Add("@ParamListaMarca", filtro.Marca == null ? "" : string.Join(",", filtro.Marca.Select(s => string.Concat(s.IdItem))));

            parametros.Add("@ParamCodUser", filtro.CodUser);
            parametros.Add("@ParamCodIdioma", filtro.CodIdioma);
            parametros.Add("@ParamSequencia", filtro.Sequencia);

            return parametros;
        }

        public DynamicParameters MontaParametrosFiltroPadraoComparativoMarcasExcel(FiltroPadraoExcel filtro, List<PadraoComboFiltro> marcas, int sequencia)
        {
            var parametros = new DynamicParameters();
            parametros.Add("@ParamListaTarget", filtro.Target[0] == null ? "" : string.Join(",", filtro.Target.Select(s => string.Concat(s.IdItem))));
            //parametros.Add("@ParamListaTarget", ""); //parametros.Add("@ParamListaTarget", filtro.Target == null ? "" : string.Join(",", filtro.Target.Select(s => string.Concat(s.IdItem))));
            //parametros.Add("@ParamListaTarget", (filtro.Target == null || filtro.Target.FirstOrDefault() == null) ? "" : string.Join(",", filtro.Target.Select(s => string.Concat(s.IdItem))));
            parametros.Add("@ParamListaRegiao", filtro.Regiao == null ? "" : string.Join(",", filtro.Regiao.Select(s => string.Concat(s.IdItem))));
            parametros.Add("@ParamListaRegiaoTipo", filtro.Regiao == null ? "" : string.Join(",", filtro.Regiao.Select(s => string.Concat(s.Tipo))));
            parametros.Add("@ParamListaDemografico", filtro.Demografico == null ? "" : string.Join(",", filtro.Demografico.Select(s => string.Concat(s.IdItem))));
            parametros.Add("@ParamListaDemograficoTipo", filtro.Demografico == null ? "" : string.Join(",", filtro.Demografico.Select(s => string.Concat(s.Tipo))));

            parametros.Add("@ParamListaOnda", filtro.Onda == null ? "" : string.Join(",", filtro.Onda.Select(s => string.Concat(s.IdItem))));
            parametros.Add("@ParamListaMarca", marcas == null ? "" : string.Join(",", marcas.Select(s => string.Concat(s.IdItem))));

            parametros.Add("@ParamCodUser", filtro.CodUser);
            parametros.Add("@ParamCodIdioma", filtro.CodIdioma);
            parametros.Add("@ParamSequencia", sequencia);


            return parametros;
        }

        public DynamicParameters MontaParametrosFiltroPadraoComparativoMarcasExcelDenominator(FiltroPadraoExcel filtro, List<PadraoComboFiltro> marcas, int sequencia)
        {
            //var parametros = new DynamicParameters();
            //parametros.Add("@ParamListaTarget", ""); //parametros.Add("@ParamListaTarget", filtro.Target == null ? "" : string.Join(",", filtro.Target.Select(s => string.Concat(s.IdItem))));
            ////parametros.Add("@ParamListaTarget", (filtro.Target == null || filtro.Target.FirstOrDefault() == null) ? "" : string.Join(",", filtro.Target.Select(s => string.Concat(s.IdItem))));
            //parametros.Add("@ParamListaRegiao", filtro.Regiao == null ? "" : string.Join(",", filtro.Regiao.Select(s => string.Concat(s.IdItem))));
            //parametros.Add("@ParamListaRegiaoTipo", filtro.Regiao == null ? "" : string.Join(",", filtro.Regiao.Select(s => string.Concat(s.Tipo))));
            //parametros.Add("@ParamListaDemografico", filtro.Demografico == null ? "" : string.Join(",", filtro.Demografico.Select(s => string.Concat(s.IdItem))));
            //parametros.Add("@ParamListaDemograficoTipo", filtro.Demografico == null ? "" : string.Join(",", filtro.Demografico.Select(s => string.Concat(s.Tipo))));

            //parametros.Add("@ParamListaOnda", filtro.Onda == null ? "" : string.Join(",", filtro.Onda.Select(s => string.Concat(s.IdItem))));
            //parametros.Add("@ParamListaMarca", marcas == null ? "" : string.Join(",", marcas.Select(s => string.Concat(s.IdItem))));

            //parametros.Add("@ParamCodUser", filtro.CodUser);
            //parametros.Add("@ParamCodIdioma", filtro.CodIdioma);
            //parametros.Add("@ParamSequencia", sequencia);
            //parametros.Add("@ParamDenominators", filtro.ParamDenominators);

            var parametros = new DynamicParameters();
            //parametros.Add("@ParamListaTarget", ""); //parametros.Add("@ParamListaTarget", filtro.Target == null ? "" : string.Join(",", filtro.Target.Select(s => string.Concat(s.IdItem))));
            parametros.Add("@ParamListaTarget", filtro.Target[0] == null ? "" : string.Join(",", filtro.Target.Select(s => string.Concat(s.IdItem))));
            parametros.Add("@ParamListaRegiao", filtro.Regiao == null ? "" : string.Join(",", filtro.Regiao.Select(s => string.Concat(s.IdItem))));
            parametros.Add("@ParamListaRegiaoTipo", filtro.Regiao == null ? "" : string.Join(",", filtro.Regiao.Select(s => string.Concat(s.Tipo))));
            parametros.Add("@ParamListaDemografico", filtro.Demografico == null ? "" : string.Join(",", filtro.Demografico.Select(s => string.Concat(s.IdItem))));
            parametros.Add("@ParamListaDemograficoTipo", filtro.Demografico == null ? "" : string.Join(",", filtro.Demografico.Select(s => string.Concat(s.Tipo))));
            parametros.Add("@ParamListaOnda", filtro.Onda == null ? "" : string.Join(",", filtro.Onda.Select(s => string.Concat(s.IdItem))));
            parametros.Add("@ParamListaMarca", marcas == null ? "" : string.Join(",", marcas.Select(s => string.Concat(s.IdItem))));

            parametros.Add("@ParamCodUser", filtro.CodUser);
            parametros.Add("@ParamCodIdioma", filtro.CodIdioma);
            parametros.Add("@ParamSequencia", sequencia);

            return parametros;
        }


        public DynamicParameters MontaParametrosFiltroPadraoComparativoMarcasExcelBIA(FiltroPadraoExcel filtro, List<PadraoComboFiltro> marcas, int sequencia)
        {
            var parametros = new DynamicParameters();
            parametros.Add("@ParamListaTarget", filtro.Target[0] == null ? "" : string.Join(",", filtro.Target.Select(s => string.Concat(s.IdItem))));
            //parametros.Add("@ParamListaTarget", (filtro.Target == null || filtro.Target.FirstOrDefault() == null) ? "" : string.Join(",", filtro.Target.Select(s => string.Concat(s.IdItem))));
            parametros.Add("@ParamListaRegiao", filtro.Regiao == null ? "" : string.Join(",", filtro.Regiao.Select(s => string.Concat(s.IdItem))));
            parametros.Add("@ParamListaRegiaoTipo", filtro.Regiao == null ? "" : string.Join(",", filtro.Regiao.Select(s => string.Concat(s.Tipo))));
            parametros.Add("@ParamListaDemografico", filtro.Demografico == null ? "" : string.Join(",", filtro.Demografico.Select(s => string.Concat(s.IdItem))));
            parametros.Add("@ParamListaDemograficoTipo", filtro.Demografico == null ? "" : string.Join(",", filtro.Demografico.Select(s => string.Concat(s.Tipo))));

            parametros.Add("@ParamListaOnda", filtro.Onda == null ? "" : string.Join(",", filtro.Onda.Select(s => string.Concat(s.IdItem))));
            parametros.Add("@ParamListaMarca", marcas == null ? "" : string.Join(",", marcas.Select(s => string.Concat(s.IdItem))));

            parametros.Add("@ParamCodUser", filtro.CodUser);
            parametros.Add("@ParamCodIdioma", filtro.CodIdioma);
            parametros.Add("@ParamSequencia", sequencia);
            parametros.Add("@ParamBIA", filtro.ParamBIA);


            return parametros;
        }

        public DynamicParameters MontaParametrosFiltroPadraoComparativoMarcasExcelNine(FiltroPadraoExcel filtro, List<PadraoComboFiltro> marcas, int sequencia)
        {
            var parametros = new DynamicParameters();
            parametros.Add("@ParamListaTarget", ""); //parametros.Add("@ParamListaTarget", filtro.Target == null ? "" : string.Join(",", filtro.Target.Select(s => string.Concat(s.IdItem))));
            //parametros.Add("@ParamListaTarget", (filtro.Target == null || filtro.Target.FirstOrDefault() == null) ? "" : string.Join(",", filtro.Target.Select(s => string.Concat(s.IdItem))));
            parametros.Add("@ParamListaRegiao", filtro.Regiao == null ? "" : string.Join(",", filtro.Regiao.Select(s => string.Concat(s.IdItem))));
            parametros.Add("@ParamListaRegiaoTipo", filtro.Regiao == null ? "" : string.Join(",", filtro.Regiao.Select(s => string.Concat(s.Tipo))));
            parametros.Add("@ParamListaDemografico", filtro.Demografico == null ? "" : string.Join(",", filtro.Demografico.Select(s => string.Concat(s.IdItem))));
            parametros.Add("@ParamListaDemograficoTipo", filtro.Demografico == null ? "" : string.Join(",", filtro.Demografico.Select(s => string.Concat(s.Tipo))));

            parametros.Add("@ParamListaOnda", filtro.Onda == null ? "" : string.Join(",", filtro.Onda.Select(s => string.Concat(s.IdItem))));
            parametros.Add("@ParamListaMarca", marcas == null ? "" : string.Join(",", marcas.Select(s => string.Concat(s.IdItem))));

            parametros.Add("@ParamCodUser", filtro.CodUser);
            parametros.Add("@ParamCodIdioma", filtro.CodIdioma);
            parametros.Add("@ParamSequencia", sequencia);

            return parametros;
        }


        public DynamicParameters MontaParametrosFiltroPadraoEvolutivoMarcasExcel(FiltroPadraoExcel filtro, List<PadraoComboFiltro> ondas, int sequencia)
        {
            var parametros = new DynamicParameters();
            parametros.Add("@ParamListaTarget", ""); //parametros.Add("@ParamListaTarget", filtro.Target == null ? "" : string.Join(",", filtro.Target.Select(s => string.Concat(s.IdItem))));
            //parametros.Add("@ParamListaTarget", (filtro.Target == null || filtro.Target.FirstOrDefault() == null) ? "" : string.Join(",", filtro.Target.Select(s => string.Concat(s.IdItem))));
            parametros.Add("@ParamListaRegiao", filtro.Regiao == null ? "" : string.Join(",", filtro.Regiao.Select(s => string.Concat(s.IdItem))));
            parametros.Add("@ParamListaRegiaoTipo", filtro.Regiao == null ? "" : string.Join(",", filtro.Regiao.Select(s => string.Concat(s.Tipo))));
            parametros.Add("@ParamListaDemografico", filtro.Demografico == null ? "" : string.Join(",", filtro.Demografico.Select(s => string.Concat(s.IdItem))));
            parametros.Add("@ParamListaDemograficoTipo", filtro.Demografico == null ? "" : string.Join(",", filtro.Demografico.Select(s => string.Concat(s.Tipo))));

            parametros.Add("@ParamListaOnda", ondas == null || ondas.FirstOrDefault() == null ? "" : string.Join(",", ondas.Select(s => string.Concat(s.IdItem))));
            parametros.Add("@ParamListaMarca", filtro.Marca == null ? "" : string.Join(",", filtro.Marca.Select(s => string.Concat(s.IdItem))));

            parametros.Add("@ParamCodUser", filtro.CodUser);
            parametros.Add("@ParamCodIdioma", filtro.CodIdioma);
            parametros.Add("@ParamSequencia", sequencia);

            return parametros;
        }

        public DynamicParameters MontaParametrosFiltroPadraoEvolutivoMarcasExcelDenominator(FiltroPadraoExcel filtro, List<PadraoComboFiltro> ondas, int sequencia)
        {
            //var parametros = new DynamicParameters();
            //parametros.Add("@ParamListaTarget", ""); //parametros.Add("@ParamListaTarget", filtro.Target == null ? "" : string.Join(",", filtro.Target.Select(s => string.Concat(s.IdItem))));
            ////parametros.Add("@ParamListaTarget", (filtro.Target == null || filtro.Target.FirstOrDefault() == null) ? "" : string.Join(",", filtro.Target.Select(s => string.Concat(s.IdItem))));
            //parametros.Add("@ParamListaRegiao", filtro.Regiao == null ? "" : string.Join(",", filtro.Regiao.Select(s => string.Concat(s.IdItem))));
            //parametros.Add("@ParamListaRegiaoTipo", filtro.Regiao == null ? "" : string.Join(",", filtro.Regiao.Select(s => string.Concat(s.Tipo))));
            //parametros.Add("@ParamListaDemografico", filtro.Demografico == null ? "" : string.Join(",", filtro.Demografico.Select(s => string.Concat(s.IdItem))));
            //parametros.Add("@ParamListaDemograficoTipo", filtro.Demografico == null ? "" : string.Join(",", filtro.Demografico.Select(s => string.Concat(s.Tipo))));

            //parametros.Add("@ParamListaOnda", ondas == null || ondas.FirstOrDefault() == null ? "" : string.Join(",", ondas.Select(s => string.Concat(s.IdItem))));
            //parametros.Add("@ParamListaMarca", filtro.Marca == null ? "" : string.Join(",", filtro.Marca.Select(s => string.Concat(s.IdItem))));

            //parametros.Add("@ParamCodUser", filtro.CodUser);
            //parametros.Add("@ParamCodIdioma", filtro.CodIdioma);
            //parametros.Add("@ParamSequencia", sequencia);
            //parametros.Add("@ParamDenominators", filtro.ParamDenominators);
            //return parametros;

            if (ondas[0] == null)
            {
                ondas = null;
            }

            var parametros = new DynamicParameters();
            //parametros.Add("@ParamListaTarget", ""); //parametros.Add("@ParamListaTarget", filtro.Target == null ? "" : string.Join(",", filtro.Target.Select(s => string.Concat(s.IdItem))));
            parametros.Add("@ParamListaTarget", filtro.Target[0] == null ? "" : string.Join(",", filtro.Target.Select(s => string.Concat(s.IdItem))));
            parametros.Add("@ParamListaRegiao", filtro.Regiao == null ? "" : string.Join(",", filtro.Regiao.Select(s => string.Concat(s.IdItem))));
            parametros.Add("@ParamListaRegiaoTipo", filtro.Regiao == null ? "" : string.Join(",", filtro.Regiao.Select(s => string.Concat(s.Tipo))));
            parametros.Add("@ParamListaDemografico", filtro.Demografico == null ? "" : string.Join(",", filtro.Demografico.Select(s => string.Concat(s.IdItem))));
            parametros.Add("@ParamListaDemograficoTipo", filtro.Demografico == null ? "" : string.Join(",", filtro.Demografico.Select(s => string.Concat(s.Tipo))));
            parametros.Add("@ParamListaOnda", ondas == null || ondas.FirstOrDefault() == null ? "" : string.Join(",", ondas.Select(s => string.Concat(s.IdItem))));
            parametros.Add("@ParamListaMarca", filtro.Marca == null ? "" : string.Join(",", filtro.Marca.Select(s => string.Concat(s.IdItem))));

            parametros.Add("@ParamCodUser", filtro.CodUser);
            parametros.Add("@ParamCodIdioma", filtro.CodIdioma);
            parametros.Add("@ParamSequencia", sequencia);
            return parametros;
        }


        public DynamicParameters MontaParametrosFiltroPadraoEvolutivoMarcasExcelTop(FiltroPadraoExcel filtro, List<PadraoComboFiltro> ondas, int sequencia)
        {
            var parametros = new DynamicParameters();
            parametros.Add("@ParamListaTarget", filtro.Target[0] == null ? "" : string.Join(",", filtro.Target.Select(s => string.Concat(s.IdItem))));
            parametros.Add("@ParamListaRegiao", filtro.Regiao == null ? "" : string.Join(",", filtro.Regiao.Select(s => string.Concat(s.IdItem))));
            parametros.Add("@ParamListaRegiaoTipo", filtro.Regiao == null ? "" : string.Join(",", filtro.Regiao.Select(s => string.Concat(s.Tipo))));
            parametros.Add("@ParamListaDemografico", filtro.Demografico == null ? "" : string.Join(",", filtro.Demografico.Select(s => string.Concat(s.IdItem))));
            parametros.Add("@ParamListaDemograficoTipo", filtro.Demografico == null ? "" : string.Join(",", filtro.Demografico.Select(s => string.Concat(s.Tipo))));

            parametros.Add("@ParamListaOnda", ondas == null || ondas.FirstOrDefault() == null ? "" : string.Join(",", ondas.Select(s => string.Concat(s.IdItem))));
            parametros.Add("@ParamListaMarca", filtro.Marca == null ? "" : string.Join(",", filtro.Marca.Select(s => string.Concat(s.IdItem))));

            parametros.Add("@ParamCodUser", filtro.CodUser);
            parametros.Add("@ParamCodIdioma", filtro.CodIdioma);
            parametros.Add("@ParamSequencia", sequencia);
            parametros.Add("@ParamTipo", filtro.ParamTipo);
            parametros.Add("@ParamOndaAtual", filtro.ParamOndaAtual);


            return parametros;
        }


        public DynamicParameters MontaParametrosFiltroPadraoEvolutivoMarcasDuploExcel(FiltroPadraoExcelDuplo filtro, PadraoComboFiltro onda, PadraoComboFiltro marca, int sequencia)
        {
            var parametros = new DynamicParameters();
            parametros.Add("@ParamListaTarget", ""); //parametros.Add("@ParamListaTarget", filtro.Target == null ? "" : string.Join(",", filtro.Target.Select(s => string.Concat(s.IdItem))));
            //parametros.Add("@ParamListaTarget", (filtro.Target == null || filtro.Target.FirstOrDefault() == null) ? "" : string.Join(",", filtro.Target.Select(s => string.Concat(s.IdItem))));
            parametros.Add("@ParamListaRegiao", filtro.Regiao == null ? "" : string.Join(",", filtro.Regiao.Select(s => string.Concat(s.IdItem))));
            parametros.Add("@ParamListaRegiaoTipo", filtro.Regiao == null ? "" : string.Join(",", filtro.Regiao.Select(s => string.Concat(s.Tipo))));
            parametros.Add("@ParamListaDemografico", filtro.Demografico == null ? "" : string.Join(",", filtro.Demografico.Select(s => string.Concat(s.IdItem))));
            parametros.Add("@ParamListaDemograficoTipo", filtro.Demografico == null ? "" : string.Join(",", filtro.Demografico.Select(s => string.Concat(s.Tipo))));

            parametros.Add("@ParamListaOnda", onda == null ? "" : string.Join(",", onda.IdItem));
            parametros.Add("@ParamListaMarca", marca == null ? "" : string.Join(",", marca.IdItem));

            parametros.Add("@ParamCodUser", filtro.CodUser);
            parametros.Add("@ParamCodIdioma", filtro.CodIdioma);
            parametros.Add("@ParamSequencia", sequencia);

            return parametros;
        }

        public DynamicParameters MontaParametrosFiltros(ParamGeralFiltro filtro)
        {
            var parametros = new DynamicParameters();
            parametros.Add("@ParamCodUser", filtro.CodUserParam);
            parametros.Add("@ParamCodIdioma", filtro.CodIdiomaParam);

            return parametros;
        }

        public DynamicParameters MontaParametrosFiltrosMarca(ParamGeralFiltro filtro)
        {
            var parametros = new DynamicParameters();
            parametros.Add("@ParamCodUser", filtro.CodUserParam);
            parametros.Add("@ParamCodIdioma", filtro.CodIdiomaParam);
            parametros.Add("@ParamListaRegiao", filtro.Regiao == null ? "" : string.Join(",", filtro.Regiao.Select(s => string.Concat(s.IdItem))));
            parametros.Add("@ParamListaRegiaoTipo", filtro.Regiao == null ? "" : string.Join(",", filtro.Regiao.Select(s => string.Concat(s.Tipo))));
            //parametros.Add("@ParamDenominators", filtro.ParamTipoChocolate);
            parametros.Add("@ParamCodOnda", filtro.CodOndaParam);
            

            return parametros;
        }


        public DynamicParameters MontaParametrosFiltroBVCEvolutivo(FiltroPadrao filtro)
        {

            if (filtro.Onda[0] == null)
            {
                filtro.Onda = null;
            }

            var parametros = new DynamicParameters();
            parametros.Add("@ParamListaTarget", filtro.Target[0] == null ? "" : string.Join(",", filtro.Target.Select(s => string.Concat(s.IdItem))));
            parametros.Add("@ParamListaRegiao", filtro.Regiao == null ? "" : string.Join(",", filtro.Regiao.Select(s => string.Concat(s.IdItem))));
            parametros.Add("@ParamListaRegiaoTipo", filtro.Regiao == null ? "" : string.Join(",", filtro.Regiao.Select(s => string.Concat(s.Tipo))));
            parametros.Add("@ParamListaDemografico", filtro.Demografico == null ? "" : string.Join(",", filtro.Demografico.Select(s => string.Concat(s.IdItem))));
            parametros.Add("@ParamListaDemograficoTipo", filtro.Demografico == null ? "" : string.Join(",", filtro.Demografico.Select(s => string.Concat(s.Tipo))));
            parametros.Add("@ParamListaOnda", filtro.Onda == null ? "" : string.Join(",", filtro.Onda.Select(s => string.Concat(s.IdItem))));
            parametros.Add("@ParamListaMarca", filtro.Marca == null ? "" : string.Join(",", filtro.Marca.Select(s => string.Concat(s.IdItem))));

            parametros.Add("@ParamCodUser", filtro.CodUser);
            parametros.Add("@ParamCodIdioma", filtro.CodIdioma);
            parametros.Add("@ParamSequencia", filtro.Sequencia);
            parametros.Add("@ParamTipo", filtro.ParamTipo);


            return parametros;
        }


        public DynamicParameters MontaParametrosFiltroSTB(ParamGeralFiltro filtro)
        {
            var parametros = new DynamicParameters();
            parametros.Add("@CodOndaParam", filtro.CodOndaParam);
            parametros.Add("@CodUserParam", filtro.CodUserParam);
            parametros.Add("@CodIdiomaParam", filtro.CodIdiomaParam);

            return parametros;
        }


        public DynamicParameters MontaParametrosFiltroComunicacaoQuadroResumo(FiltroPadrao filtro)
        {
            var parametros = new DynamicParameters();
            parametros.Add("@ParamListaTarget", ""); //parametros.Add("@ParamListaTarget", filtro.Target == null ? "" : string.Join(",", filtro.Target.Select(s => string.Concat(s.IdItem))));
            // parametros.Add("@ParamListaTarget", (filtro.Target == null || filtro.Target.FirstOrDefault() == null) ? "" : string.Join(",", filtro.Target.Select(s => string.Concat(s.IdItem))));
            parametros.Add("@ParamListaRegiao", filtro.Regiao == null ? "" : string.Join(",", filtro.Regiao.Select(s => string.Concat(s.IdItem))));
            parametros.Add("@ParamListaRegiaoTipo", filtro.Regiao == null ? "" : string.Join(",", filtro.Regiao.Select(s => string.Concat(s.Tipo))));
            parametros.Add("@ParamListaDemografico", filtro.Demografico == null ? "" : string.Join(",", filtro.Demografico.Select(s => string.Concat(s.IdItem))));
            parametros.Add("@ParamListaDemograficoTipo", filtro.Demografico == null ? "" : string.Join(",", filtro.Demografico.Select(s => string.Concat(s.Tipo))));
            parametros.Add("@ParamListaOnda", filtro.Onda == null ? "" : string.Join(",", filtro.Onda.Select(s => string.Concat(s.IdItem))));
            parametros.Add("@ParamListaMarca", filtro.Marca == null ? "" : string.Join(",", filtro.Marca.Select(s => string.Concat(s.IdItem))));

            parametros.Add("@ParamCodUser", filtro.CodUser);
            parametros.Add("@ParamCodIdioma", filtro.CodIdioma);
            parametros.Add("@ParamSequencia", filtro.Sequencia);
            parametros.Add("@ParamSTB", filtro.ParamSTB);


            return parametros;
        }

        public DynamicParameters MontaParametrosFiltroPadraoNine(FiltroPadrao filtro)
        {
            var parametros = new DynamicParameters();
            parametros.Add("@ParamListaTarget", ""); //parametros.Add("@ParamListaTarget", filtro.Target == null ? "" : string.Join(",", filtro.Target.Select(s => string.Concat(s.IdItem))));
            //parametros.Add("@ParamListaTarget", (filtro.Target == null || filtro.Target.FirstOrDefault() == null)  ? "" : string.Join(",", filtro.Target.Select(s => string.Concat(s.IdItem))));
            parametros.Add("@ParamListaRegiao", filtro.Regiao == null ? "" : string.Join(",", filtro.Regiao.Select(s => string.Concat(s.IdItem))));
            parametros.Add("@ParamListaRegiaoTipo", filtro.Regiao == null ? "" : string.Join(",", filtro.Regiao.Select(s => string.Concat(s.Tipo))));
            parametros.Add("@ParamListaDemografico", filtro.Demografico == null ? "" : string.Join(",", filtro.Demografico.Select(s => string.Concat(s.IdItem))));
            parametros.Add("@ParamListaDemograficoTipo", filtro.Demografico == null ? "" : string.Join(",", filtro.Demografico.Select(s => string.Concat(s.Tipo))));
            parametros.Add("@ParamListaOnda", filtro.Onda == null ? "" : string.Join(",", filtro.Onda.Select(s => string.Concat(s.IdItem))));
            parametros.Add("@ParamListaMarca", filtro.Marca == null ? "" : string.Join(",", filtro.Marca.Select(s => string.Concat(s.IdItem))));

            parametros.Add("@ParamCodUser", filtro.CodUser);
            parametros.Add("@ParamCodIdioma", filtro.CodIdioma);
            parametros.Add("@ParamSequencia", filtro.Sequencia);

            return parametros;
        }

    }
}
