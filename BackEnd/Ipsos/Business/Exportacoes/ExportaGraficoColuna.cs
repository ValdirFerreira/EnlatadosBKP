using Entities.Filtros;
using Entities.GraficoAranha;
using Entities.GraficoBVCEvolutivo;
using Entities.GraficoBVCTop10;
using Entities.GraficoColunas;
using Entities.GraficoComunicacaoDiagnostico;
using Entities.GraficoComunicacaoRecall;
using Entities.GraficoComunicacaoVisto;
using Entities.GraficoFunil;
using Entities.GraficoImagemEvolutivo;
using Entities.GraficoImagemPosicionamento;
using Entities.GraficoLinhasImagem;
using Entities.Parametros;
using Entities.TabelaAdHoc;
using Entities.TraducaoIdioma;
using Helpers.ExportacaoExcel;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Web;

namespace Business.Exportacoes
{
    public class ExportaGraficoColuna
    {
        private string AdjustPath(string path)
        {
            if (path[path.Length - 1] == '\\')
                return path;
            else
                return path + @"\";
        }

        private void GetNewFileMemory(ref string pathOrigem)
        {
            try
            {
                string pathDestination = AdjustPath(HttpContext.Current.Server.MapPath("~")) + @"arquivos\excelExportacao\workspace\" + "package_template_temp_" + Guid.NewGuid().ToString() + ".xlsx";

                using (var file = File.OpenRead(pathOrigem))
                {
                    using (var filedestino = File.Create(pathDestination))
                    {
                        file.CopyTo(filedestino);
                    }
                }

                pathOrigem = pathDestination;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        #region DownloadDashboardTwo
        public byte[] DownloadDashboardTwoComparativoMarcas(GraficoColunasFullLoad grafico, string tituloPlanilha, List<TraducaoComponente> listaTraducaoComponente)
        {
            string fileName = "";
            string templateFileName = "";
            try
            {
                string strDownloadNome;

                strDownloadNome = "Modelo_Exportação_Gráficos_Padrao";

                string strPath = AdjustPath(HttpContext.Current.Server.MapPath("~")) + @"Arquivos\ExcelExportacao\workspace\";
                fileName = strPath + strDownloadNome + DateTime.Now.ToString("yyyy - MM - dd HH':'mm':'ss':'fff").Replace(" - ", "").Replace(":", "").Replace(" ", "") + ".xlsx";

                templateFileName = AdjustPath(HttpContext.Current.Server.MapPath("~")) + @"Arquivos\ExcelExportacao\Modelos\" + strDownloadNome + ".xlsx";

                GetNewFileMemory(ref templateFileName);

                if (File.Exists(fileName))
                    File.Delete(fileName);

                MemoryStream returnedMemoryStream = null;

                using (var file = new FileStream(fileName, FileMode.CreateNew))
                {
                    using (var temp = new FileStream(templateFileName, FileMode.Open))
                    {
                        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                        ExcelPackage package = new ExcelPackage(file, temp);

                        ExcelWorksheet wworksheet = package.Workbook.Worksheets["Data"];

                        wworksheet.View.ShowHeaders = true;
                        wworksheet.View.ShowGridLines = false;
                        //Titulo da planilha. Deixar linha 3 vazias formatadas no modelo excel.
                        wworksheet.Cells[3, 2].Value = tituloPlanilha;

                        int currentRow = 6;
                        int colNameTable = 3;
                        int lastColName = 0;


                        if (grafico != null)
                        {
                            // Preenche cabeçalho com os periodos:
                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoColunas1.DescMarca;
                            colNameTable++;
                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoColunas2.DescMarca;
                            colNameTable++;

                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoColunas3.DescMarca;
                            colNameTable++;

                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoColunas4.DescMarca;
                            colNameTable++;

                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoColunas5.DescMarca;
                            colNameTable++;

                            colNameTable--;
                            lastColName = colNameTable;

                            //Formata cabeçalho
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(3, 6) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Font.Size = 11;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(3, 6) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Font.Name = "Arial Nova Cond";
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(3, 6) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Font.Bold = true;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(3, 6) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(3, 6) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(3, 6) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(3, 6) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(3, 6) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(3, 6) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                            colNameTable = 2;
                            currentRow++;


                            ////Preenche lateral
                            //wworksheet.Cells[currentRow, colNameTable].Value = "Total Awareness";
                            //currentRow++;
                            //wworksheet.Cells[currentRow, colNameTable].Value = "Prompeted";
                            //currentRow++;
                            //wworksheet.Cells[currentRow, colNameTable].Value = "Spontaneous - OM";
                            //currentRow++;
                            //wworksheet.Cells[currentRow, colNameTable].Value = "Awareness - TOM";
                            //currentRow++;
                            //wworksheet.Cells[currentRow, colNameTable].Value = "Base";

                            wworksheet.Cells[currentRow, colNameTable].Value = listaTraducaoComponente.FirstOrDefault(a => a.Objeto.Equals("dashboard-two-leg-total-Awareness")).Texto;
                            currentRow++;
                            wworksheet.Cells[currentRow, colNameTable].Value = listaTraducaoComponente.FirstOrDefault(a => a.Objeto.Equals("dashboard-two-leg-Prompeted")).Texto;
                            currentRow++;
                            wworksheet.Cells[currentRow, colNameTable].Value = listaTraducaoComponente.FirstOrDefault(a => a.Objeto.Equals("dashboard-two-leg-Spontaneous")).Texto;
                            currentRow++;
                            wworksheet.Cells[currentRow, colNameTable].Value = listaTraducaoComponente.FirstOrDefault(a => a.Objeto.Equals("dashboard-two-leg-Awareness")).Texto;
                            currentRow++;
                            wworksheet.Cells[currentRow, colNameTable].Value = "Base";

                            currentRow--;
                            var lastRow = currentRow;
                            colNameTable = 3;
                            currentRow = 7;



                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoColunas1.PercTotal;
                            if (grafico.GraficoColunas1.TesteSigTotal == "MENOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Red);
                            if (grafico.GraficoColunas1.TesteSigTotal == "MAIOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Blue);

                            currentRow++;
                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoColunas1.PercPrompeted;
                            if (grafico.GraficoColunas1.TesteSIGPrompeted == "MENOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Red);
                            if (grafico.GraficoColunas1.TesteSIGPrompeted == "MAIOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Blue);

                            currentRow++;
                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoColunas1.PercOM;
                            if (grafico.GraficoColunas1.TesteSIGOM == "MENOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Red);
                            if (grafico.GraficoColunas1.TesteSIGOM == "MAIOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Blue);

                            currentRow++;
                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoColunas1.PercTOM;
                            if (grafico.GraficoColunas1.TesteSIGTOM == "MENOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Red);
                            if (grafico.GraficoColunas1.TesteSIGTOM == "MAIOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Blue);
                            currentRow++;

                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoColunas1.BaseAbs;
                            colNameTable++;
                            currentRow = 7;

                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoColunas2.PercTotal;
                            if (grafico.GraficoColunas2.TesteSigTotal == "MENOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Red);
                            if (grafico.GraficoColunas2.TesteSigTotal == "MAIOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Blue);
                            currentRow++;

                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoColunas2.PercPrompeted;
                            if (grafico.GraficoColunas2.TesteSIGPrompeted == "MENOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Red);
                            if (grafico.GraficoColunas2.TesteSIGPrompeted == "MAIOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Blue);
                            currentRow++;

                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoColunas2.PercOM;
                            if (grafico.GraficoColunas2.TesteSIGOM == "MENOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Red);
                            if (grafico.GraficoColunas2.TesteSIGOM == "MAIOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Blue);
                            currentRow++;

                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoColunas2.PercTOM;
                            if (grafico.GraficoColunas2.TesteSIGTOM == "MENOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Red);
                            if (grafico.GraficoColunas2.TesteSIGTOM == "MAIOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Blue);
                            currentRow++;

                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoColunas2.BaseAbs;
                            colNameTable++;
                            currentRow = 7;

                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoColunas3.PercTotal;
                            if (grafico.GraficoColunas3.TesteSigTotal == "MENOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Red);
                            if (grafico.GraficoColunas3.TesteSigTotal == "MAIOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Blue);
                            currentRow++;

                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoColunas3.PercPrompeted;
                            if (grafico.GraficoColunas3.TesteSIGPrompeted == "MENOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Red);
                            if (grafico.GraficoColunas3.TesteSIGPrompeted == "MAIOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Blue);
                            currentRow++;

                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoColunas3.PercOM;
                            if (grafico.GraficoColunas3.TesteSIGOM == "MENOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Red);
                            if (grafico.GraficoColunas3.TesteSIGOM == "MAIOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Blue);
                            currentRow++;

                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoColunas3.PercTOM;
                            if (grafico.GraficoColunas3.TesteSIGTOM == "MENOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Red);
                            if (grafico.GraficoColunas3.TesteSIGTOM == "MAIOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Blue);
                            currentRow++;

                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoColunas3.BaseAbs;
                            colNameTable++;
                            currentRow = 7;

                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoColunas4.PercTotal;
                            if (grafico.GraficoColunas4.TesteSigTotal == "MENOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Red);
                            if (grafico.GraficoColunas4.TesteSigTotal == "MAIOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Blue);
                            currentRow++;

                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoColunas4.PercPrompeted;
                            if (grafico.GraficoColunas4.TesteSIGPrompeted == "MENOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Red);
                            if (grafico.GraficoColunas4.TesteSIGPrompeted == "MAIOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Blue);
                            currentRow++;

                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoColunas4.PercOM;
                            if (grafico.GraficoColunas4.TesteSIGOM == "MENOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Red);
                            if (grafico.GraficoColunas4.TesteSIGOM == "MAIOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Blue);
                            currentRow++;

                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoColunas4.PercTOM;
                            if (grafico.GraficoColunas4.TesteSIGTOM == "MENOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Red);
                            if (grafico.GraficoColunas4.TesteSIGTOM == "MAIOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Blue);
                            currentRow++;

                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoColunas4.BaseAbs;
                            colNameTable++;
                            currentRow = 7;

                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoColunas5.PercTotal;
                            if (grafico.GraficoColunas5.TesteSigTotal == "MENOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Red);
                            if (grafico.GraficoColunas5.TesteSigTotal == "MAIOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Blue);
                            currentRow++;

                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoColunas5.PercPrompeted;
                            if (grafico.GraficoColunas5.TesteSIGPrompeted == "MENOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Red);
                            if (grafico.GraficoColunas5.TesteSIGPrompeted == "MAIOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Blue);
                            currentRow++;

                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoColunas5.PercOM;
                            if (grafico.GraficoColunas5.TesteSIGOM == "MENOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Red);
                            if (grafico.GraficoColunas5.TesteSIGOM == "MAIOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Blue);
                            currentRow++;

                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoColunas5.PercTOM;
                            if (grafico.GraficoColunas5.TesteSIGTOM == "MENOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Red);
                            if (grafico.GraficoColunas5.TesteSIGTOM == "MAIOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Blue);
                            currentRow++;

                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoColunas5.BaseAbs;
                            colNameTable++;
                            currentRow = 7;


                            lastRow++;
                            currentRow = lastRow;
                            colNameTable--;
                            lastColName = colNameTable;

                            //Formata dados
                            //wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 7) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Numberformat.Format = "0.00";
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 7) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Font.Size = 11;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 7) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Font.Name = "Arial Nova Cond";
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 7) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Font.Bold = true;

                            // Preenche lateral
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 11) + ":" + ExcelUtil.GetCelulaByColumnIndex(3, 11)].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 11) + ":" + ExcelUtil.GetCelulaByColumnIndex(3, 11)].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 11) + ":" + ExcelUtil.GetCelulaByColumnIndex(3, 11)].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 11) + ":" + ExcelUtil.GetCelulaByColumnIndex(3, 11)].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 11) + ":" + ExcelUtil.GetCelulaByColumnIndex(3, 11)].Style.Border.Top.Style = ExcelBorderStyle.Thin;

                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 7) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 7) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 7) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 7) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 7) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                        }
                        wworksheet.Select();

                        using (MemoryStream memoryStream = new MemoryStream())
                        {
                            package.SaveAs(memoryStream);

                            memoryStream.Position = 0;

                            returnedMemoryStream = memoryStream;
                        }
                    }
                }

                return File.ReadAllBytes(@"" + fileName);
            }

            catch (Exception)
            {
                throw;
            }
            finally
            {
                File.Delete(fileName);
                File.Delete(templateFileName);
            }
        }

        public byte[] DownloadDashboardTwoEvolutivoMarcas(GraficoColunasFullLoad grafico, string tituloPlanilha, List<TraducaoComponente> listaTraducaoComponente, FiltroPadraoExcel filtro)
        {
            string fileName = "";
            string templateFileName = "";
            try
            {
                string strDownloadNome;

                strDownloadNome = "Modelo_Exportação_Gráficos_Padrao";

                string strPath = AdjustPath(HttpContext.Current.Server.MapPath("~")) + @"Arquivos\ExcelExportacao\workspace\";
                fileName = strPath + strDownloadNome + DateTime.Now.ToString("yyyy - MM - dd HH':'mm':'ss':'fff").Replace(" - ", "").Replace(":", "").Replace(" ", "") + ".xlsx";

                templateFileName = AdjustPath(HttpContext.Current.Server.MapPath("~")) + @"Arquivos\ExcelExportacao\Modelos\" + strDownloadNome + ".xlsx";

                GetNewFileMemory(ref templateFileName);

                if (File.Exists(fileName))
                    File.Delete(fileName);

                MemoryStream returnedMemoryStream = null;

                using (var file = new FileStream(fileName, FileMode.CreateNew))
                {
                    using (var temp = new FileStream(templateFileName, FileMode.Open))
                    {
                        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                        ExcelPackage package = new ExcelPackage(file, temp);

                        ExcelWorksheet wworksheet = package.Workbook.Worksheets["Data"];

                        wworksheet.View.ShowHeaders = true;
                        wworksheet.View.ShowGridLines = false;
                        //Titulo da planilha. Deixar linha 3 vazias formatadas no modelo excel.
                        wworksheet.Cells[3, 2].Value = tituloPlanilha;

                        int currentRow = 6;
                        int colNameTable = 3;
                        int lastColName = 0;


                        if (grafico != null)
                        {
                            // Preenche cabeçalho com os periodos:
                            if (filtro.Onda1.FirstOrDefault() != null)
                            {
                                wworksheet.Cells[currentRow, colNameTable].Value = filtro.Onda1.FirstOrDefault().DescItem;
                            }
                            else
                            {
                                wworksheet.Cells[currentRow, colNameTable].Value = "N/A";
                            }
                            colNameTable++;


                            if (filtro.Onda2.FirstOrDefault() != null)
                                wworksheet.Cells[currentRow, colNameTable].Value = filtro.Onda2.FirstOrDefault().DescItem;
                            else
                                wworksheet.Cells[currentRow, colNameTable].Value = "N/A";
                            colNameTable++;

                            if (filtro.Onda3.FirstOrDefault() == null)
                                wworksheet.Cells[currentRow, colNameTable].Value = "N/A";
                            else wworksheet.Cells[currentRow, colNameTable].Value = filtro.Onda3.FirstOrDefault().DescItem;

                            colNameTable++;

                            if (filtro.Onda4.FirstOrDefault() == null)
                                wworksheet.Cells[currentRow, colNameTable].Value = "N/A";
                            else
                                wworksheet.Cells[currentRow, colNameTable].Value = filtro.Onda4.FirstOrDefault().DescItem;

                            colNameTable++;

                            //wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoColunas5.DescMarca;
                            //colNameTable++;

                            colNameTable--;
                            lastColName = colNameTable;

                            //Formata cabeçalho
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(3, 6) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Font.Size = 11;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(3, 6) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Font.Name = "Arial Nova Cond";
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(3, 6) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Font.Bold = true;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(3, 6) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(3, 6) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(3, 6) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(3, 6) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(3, 6) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(3, 6) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                            colNameTable = 2;
                            currentRow++;


                            ////Preenche lateral
                            //wworksheet.Cells[currentRow, colNameTable].Value = "Total Awareness";
                            //currentRow++;
                            //wworksheet.Cells[currentRow, colNameTable].Value = "Prompeted";
                            //currentRow++;
                            //wworksheet.Cells[currentRow, colNameTable].Value = "Spontaneous - OM";
                            //currentRow++;
                            //wworksheet.Cells[currentRow, colNameTable].Value = "Awareness - TOM";
                            //currentRow++;
                            //wworksheet.Cells[currentRow, colNameTable].Value = "Base";

                            wworksheet.Cells[currentRow, colNameTable].Value = listaTraducaoComponente.FirstOrDefault(a => a.Objeto.Equals("dashboard-two-leg-total-Awareness")).Texto;
                            currentRow++;
                            wworksheet.Cells[currentRow, colNameTable].Value = listaTraducaoComponente.FirstOrDefault(a => a.Objeto.Equals("dashboard-two-leg-Prompeted")).Texto;
                            currentRow++;
                            wworksheet.Cells[currentRow, colNameTable].Value = listaTraducaoComponente.FirstOrDefault(a => a.Objeto.Equals("dashboard-two-leg-Spontaneous")).Texto;
                            currentRow++;
                            wworksheet.Cells[currentRow, colNameTable].Value = listaTraducaoComponente.FirstOrDefault(a => a.Objeto.Equals("dashboard-two-leg-Awareness")).Texto;
                            currentRow++;
                            wworksheet.Cells[currentRow, colNameTable].Value = "Base";

                            currentRow--;
                            var lastRow = currentRow;
                            colNameTable = 3;
                            currentRow = 7;






                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoColunas1.PercTotal;
                            if (grafico.GraficoColunas1.TesteSigTotal == "MENOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Red);
                            if (grafico.GraficoColunas1.TesteSigTotal == "MAIOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Blue);
                            currentRow++;

                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoColunas1.PercPrompeted;
                            if (grafico.GraficoColunas1.TesteSIGPrompeted == "MENOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Red);
                            if (grafico.GraficoColunas1.TesteSIGPrompeted == "MAIOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Blue);
                            currentRow++;

                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoColunas1.PercOM;
                            if (grafico.GraficoColunas1.TesteSIGOM == "MENOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Red);
                            if (grafico.GraficoColunas1.TesteSIGOM == "MAIOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Blue);
                            currentRow++;

                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoColunas1.PercTOM;
                            if (grafico.GraficoColunas1.TesteSIGTOM == "MENOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Red);
                            if (grafico.GraficoColunas1.TesteSIGTOM == "MAIOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Blue);
                            currentRow++;

                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoColunas1.BaseAbs;
                            colNameTable++;
                            currentRow = 7;

                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoColunas2.PercTotal;
                            if (grafico.GraficoColunas2.TesteSigTotal == "MENOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Red);
                            if (grafico.GraficoColunas2.TesteSigTotal == "MAIOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Blue);
                            currentRow++;

                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoColunas2.PercPrompeted;
                            if (grafico.GraficoColunas2.TesteSIGPrompeted == "MENOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Red);
                            if (grafico.GraficoColunas2.TesteSIGPrompeted == "MAIOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Blue);
                            currentRow++;

                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoColunas2.PercOM;
                            if (grafico.GraficoColunas2.TesteSIGOM == "MENOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Red);
                            if (grafico.GraficoColunas2.TesteSIGOM == "MAIOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Blue);
                            currentRow++;

                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoColunas2.PercTOM;
                            if (grafico.GraficoColunas2.TesteSIGTOM == "MENOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Red);
                            if (grafico.GraficoColunas2.TesteSIGTOM == "MAIOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Blue);
                            currentRow++;

                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoColunas2.BaseAbs;
                            colNameTable++;
                            currentRow = 7;

                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoColunas3.PercTotal;
                            if (grafico.GraficoColunas3.TesteSigTotal == "MENOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Red);
                            if (grafico.GraficoColunas3.TesteSigTotal == "MAIOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Blue);
                            currentRow++;

                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoColunas3.PercPrompeted;
                            if (grafico.GraficoColunas3.TesteSIGPrompeted == "MENOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Red);
                            if (grafico.GraficoColunas3.TesteSIGPrompeted == "MAIOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Blue);
                            currentRow++;

                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoColunas3.PercOM;
                            if (grafico.GraficoColunas3.TesteSIGOM == "MENOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Red);
                            if (grafico.GraficoColunas3.TesteSIGOM == "MAIOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Blue);
                            currentRow++;

                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoColunas3.PercTOM;
                            if (grafico.GraficoColunas3.TesteSIGTOM == "MENOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Red);
                            if (grafico.GraficoColunas3.TesteSIGTOM == "MAIOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Blue);
                            currentRow++;

                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoColunas3.BaseAbs;
                            colNameTable++;
                            currentRow = 7;


                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoColunas4.PercTotal;
                            if (grafico.GraficoColunas4.TesteSigTotal == "MENOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Red);
                            if (grafico.GraficoColunas4.TesteSigTotal == "MAIOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Blue);
                            currentRow++;

                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoColunas4.PercPrompeted;
                            if (grafico.GraficoColunas4.TesteSIGPrompeted == "MENOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Red);
                            if (grafico.GraficoColunas4.TesteSIGPrompeted == "MAIOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Blue);
                            currentRow++;

                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoColunas4.PercOM;
                            if (grafico.GraficoColunas4.TesteSIGOM == "MENOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Red);
                            if (grafico.GraficoColunas4.TesteSIGOM == "MAIOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Blue);
                            currentRow++;

                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoColunas4.PercTOM;
                            if (grafico.GraficoColunas4.TesteSIGTOM == "MENOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Red);
                            if (grafico.GraficoColunas4.TesteSIGTOM == "MAIOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Blue);
                            currentRow++;

                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoColunas4.BaseAbs;
                            colNameTable++;
                            currentRow = 7;




                            //wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoColunas5.PercTotal;
                            //currentRow++;
                            //wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoColunas5.PercPrompeted;
                            //currentRow++;
                            //wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoColunas5.PercOM;
                            //currentRow++;
                            //wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoColunas5.PercTOM;
                            //currentRow++;
                            //wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoColunas5.BaseAbs;
                            //colNameTable++;
                            //currentRow = 7;


                            lastRow++;
                            currentRow = lastRow;
                            colNameTable--;
                            lastColName = colNameTable;

                            //Formata dados
                            //wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 7) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Numberformat.Format = "0.00";
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 7) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Font.Size = 11;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 7) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Font.Name = "Arial Nova Cond";
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 7) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Font.Bold = true;

                            // Preenche lateral
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 11) + ":" + ExcelUtil.GetCelulaByColumnIndex(3, 11)].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 11) + ":" + ExcelUtil.GetCelulaByColumnIndex(3, 11)].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 11) + ":" + ExcelUtil.GetCelulaByColumnIndex(3, 11)].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 11) + ":" + ExcelUtil.GetCelulaByColumnIndex(3, 11)].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 11) + ":" + ExcelUtil.GetCelulaByColumnIndex(3, 11)].Style.Border.Top.Style = ExcelBorderStyle.Thin;

                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 7) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 7) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 7) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 7) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 7) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                        }
                        wworksheet.Select();

                        using (MemoryStream memoryStream = new MemoryStream())
                        {
                            package.SaveAs(memoryStream);

                            memoryStream.Position = 0;

                            returnedMemoryStream = memoryStream;
                        }
                    }
                }

                return File.ReadAllBytes(@"" + fileName);
            }

            catch (Exception)
            {
                throw;
            }
            finally
            {
                File.Delete(fileName);
                File.Delete(templateFileName);
            }
        }


        public byte[] DownloadDashboardTwoComparativoMarcasDenominators(GraficoColunasFullLoad grafico, string tituloPlanilha, List<TraducaoComponente> listaTraducaoComponente)
        {
            string fileName = "";
            string templateFileName = "";
            try
            {
                string strDownloadNome;

                strDownloadNome = "Modelo_Exportação_Gráficos_PadraoDenominators";

                string strPath = AdjustPath(HttpContext.Current.Server.MapPath("~")) + @"Arquivos\ExcelExportacao\workspace\";
                fileName = strPath + strDownloadNome + DateTime.Now.ToString("yyyy - MM - dd HH':'mm':'ss':'fff").Replace(" - ", "").Replace(":", "").Replace(" ", "") + ".xlsx";

                templateFileName = AdjustPath(HttpContext.Current.Server.MapPath("~")) + @"Arquivos\ExcelExportacao\Modelos\" + strDownloadNome + ".xlsx";

                GetNewFileMemory(ref templateFileName);

                if (File.Exists(fileName))
                    File.Delete(fileName);

                MemoryStream returnedMemoryStream = null;

                using (var file = new FileStream(fileName, FileMode.CreateNew))
                {
                    using (var temp = new FileStream(templateFileName, FileMode.Open))
                    {
                        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                        ExcelPackage package = new ExcelPackage(file, temp);

                        ExcelWorksheet wworksheet = package.Workbook.Worksheets["Data"];

                        wworksheet.View.ShowHeaders = true;
                        wworksheet.View.ShowGridLines = false;
                        //Titulo da planilha. Deixar linha 3 vazias formatadas no modelo excel.
                        wworksheet.Cells[3, 2].Value = tituloPlanilha;

                        int currentRow = 6;
                        int colNameTable = 3;
                        int lastColName = 0;


                        if (grafico != null)
                        {
                            // Preenche cabeçalho com os periodos:
                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoColunas1.DescMarca;
                            colNameTable++;
                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoColunas2.DescMarca;
                            colNameTable++;

                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoColunas3.DescMarca;
                            colNameTable++;

                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoColunas4.DescMarca;
                            colNameTable++;

                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoColunas5.DescMarca;
                            colNameTable++;

                            colNameTable--;
                            lastColName = colNameTable;

                            //Formata cabeçalho
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(3, 6) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Font.Size = 11;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(3, 6) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Font.Name = "Arial Nova Cond";
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(3, 6) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Font.Bold = true;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(3, 6) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(3, 6) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(3, 6) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(3, 6) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(3, 6) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(3, 6) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                            colNameTable = 2;
                            currentRow++;


                            ////Preenche lateral
                            //wworksheet.Cells[currentRow, colNameTable].Value = "Total Awareness";
                            //currentRow++;
                            //wworksheet.Cells[currentRow, colNameTable].Value = "Prompeted";
                            //currentRow++;
                            //wworksheet.Cells[currentRow, colNameTable].Value = "Spontaneous - OM";
                            //currentRow++;
                            //wworksheet.Cells[currentRow, colNameTable].Value = "Awareness - TOM";
                            //currentRow++;
                            //wworksheet.Cells[currentRow, colNameTable].Value = "Base";


                            wworksheet.Cells[currentRow, colNameTable].Value = listaTraducaoComponente.FirstOrDefault(a => a.Objeto.Equals("dashboard-two-leg-Awareness-total-2")).Texto;
                            currentRow++;
                            wworksheet.Cells[currentRow, colNameTable].Value = listaTraducaoComponente.FirstOrDefault(a => a.Objeto.Equals("dashboard-two-leg-Awareness-total-1")).Texto;
                            currentRow++;
                            wworksheet.Cells[currentRow, colNameTable].Value = "Base";

                            currentRow--;
                            var lastRow = currentRow;
                            colNameTable = 3;
                            currentRow = 7;



                            //wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoColunas1.PercTotal;
                            //if (grafico.GraficoColunas1.TesteSigTotal == "MENOR")
                            //    wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Red);
                            //if (grafico.GraficoColunas1.TesteSigTotal == "MAIOR")
                            //    wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Blue);

                            //currentRow++;
                            //wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoColunas1.PercPrompeted;
                            //if (grafico.GraficoColunas1.TesteSIGPrompeted == "MENOR")
                            //    wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Red);
                            //if (grafico.GraficoColunas1.TesteSIGPrompeted == "MAIOR")
                            //    wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Blue);

                            //currentRow++;
                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoColunas1.PercOM;
                            if (grafico.GraficoColunas1.TesteSIGOM == "MENOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Red);
                            if (grafico.GraficoColunas1.TesteSIGOM == "MAIOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Blue);

                            currentRow++;
                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoColunas1.PercTOM;
                            if (grafico.GraficoColunas1.TesteSIGTOM == "MENOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Red);
                            if (grafico.GraficoColunas1.TesteSIGTOM == "MAIOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Blue);
                            currentRow++;

                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoColunas1.BaseAbs;
                            colNameTable++;
                            currentRow = 7;

                            //wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoColunas2.PercTotal;
                            //if (grafico.GraficoColunas2.TesteSigTotal == "MENOR")
                            //    wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Red);
                            //if (grafico.GraficoColunas2.TesteSigTotal == "MAIOR")
                            //    wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Blue);
                            //currentRow++;

                            //wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoColunas2.PercPrompeted;
                            //if (grafico.GraficoColunas2.TesteSIGPrompeted == "MENOR")
                            //    wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Red);
                            //if (grafico.GraficoColunas2.TesteSIGPrompeted == "MAIOR")
                            //    wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Blue);
                            //currentRow++;

                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoColunas2.PercOM;
                            if (grafico.GraficoColunas2.TesteSIGOM == "MENOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Red);
                            if (grafico.GraficoColunas2.TesteSIGOM == "MAIOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Blue);
                            currentRow++;

                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoColunas2.PercTOM;
                            if (grafico.GraficoColunas2.TesteSIGTOM == "MENOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Red);
                            if (grafico.GraficoColunas2.TesteSIGTOM == "MAIOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Blue);
                            currentRow++;

                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoColunas2.BaseAbs;
                            colNameTable++;
                            currentRow = 7;

                            //wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoColunas3.PercTotal;
                            //if (grafico.GraficoColunas3.TesteSigTotal == "MENOR")
                            //    wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Red);
                            //if (grafico.GraficoColunas3.TesteSigTotal == "MAIOR")
                            //    wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Blue);
                            //currentRow++;

                            //wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoColunas3.PercPrompeted;
                            //if (grafico.GraficoColunas3.TesteSIGPrompeted == "MENOR")
                            //    wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Red);
                            //if (grafico.GraficoColunas3.TesteSIGPrompeted == "MAIOR")
                            //    wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Blue);
                            //currentRow++;

                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoColunas3.PercOM;
                            if (grafico.GraficoColunas3.TesteSIGOM == "MENOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Red);
                            if (grafico.GraficoColunas3.TesteSIGOM == "MAIOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Blue);
                            currentRow++;

                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoColunas3.PercTOM;
                            if (grafico.GraficoColunas3.TesteSIGTOM == "MENOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Red);
                            if (grafico.GraficoColunas3.TesteSIGTOM == "MAIOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Blue);
                            currentRow++;

                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoColunas3.BaseAbs;
                            colNameTable++;
                            currentRow = 7;

                            //wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoColunas4.PercTotal;
                            //if (grafico.GraficoColunas4.TesteSigTotal == "MENOR")
                            //    wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Red);
                            //if (grafico.GraficoColunas4.TesteSigTotal == "MAIOR")
                            //    wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Blue);
                            //currentRow++;

                            //wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoColunas4.PercPrompeted;
                            //if (grafico.GraficoColunas4.TesteSIGPrompeted == "MENOR")
                            //    wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Red);
                            //if (grafico.GraficoColunas4.TesteSIGPrompeted == "MAIOR")
                            //    wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Blue);
                            //currentRow++;

                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoColunas4.PercOM;
                            if (grafico.GraficoColunas4.TesteSIGOM == "MENOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Red);
                            if (grafico.GraficoColunas4.TesteSIGOM == "MAIOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Blue);
                            currentRow++;

                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoColunas4.PercTOM;
                            if (grafico.GraficoColunas4.TesteSIGTOM == "MENOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Red);
                            if (grafico.GraficoColunas4.TesteSIGTOM == "MAIOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Blue);
                            currentRow++;

                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoColunas4.BaseAbs;
                            colNameTable++;
                            currentRow = 7;

                            //wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoColunas5.PercTotal;
                            //if (grafico.GraficoColunas5.TesteSigTotal == "MENOR")
                            //    wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Red);
                            //if (grafico.GraficoColunas5.TesteSigTotal == "MAIOR")
                            //    wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Blue);
                            //currentRow++;

                            //wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoColunas5.PercPrompeted;
                            //if (grafico.GraficoColunas5.TesteSIGPrompeted == "MENOR")
                            //    wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Red);
                            //if (grafico.GraficoColunas5.TesteSIGPrompeted == "MAIOR")
                            //    wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Blue);
                            //currentRow++;

                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoColunas5.PercOM;
                            if (grafico.GraficoColunas5.TesteSIGOM == "MENOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Red);
                            if (grafico.GraficoColunas5.TesteSIGOM == "MAIOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Blue);
                            currentRow++;

                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoColunas5.PercTOM;
                            if (grafico.GraficoColunas5.TesteSIGTOM == "MENOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Red);
                            if (grafico.GraficoColunas5.TesteSIGTOM == "MAIOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Blue);
                            currentRow++;

                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoColunas5.BaseAbs;
                            colNameTable++;
                            currentRow = 7;


                            lastRow++;
                            currentRow = lastRow;
                            colNameTable--;
                            lastColName = colNameTable;

                            //Formata dados
                            //wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 7) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Numberformat.Format = "0.00";
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 7) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Font.Size = 11;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 7) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Font.Name = "Arial Nova Cond";
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 7) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Font.Bold = true;

                            // Preenche lateral
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 9) + ":" + ExcelUtil.GetCelulaByColumnIndex(3, 9)].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 9) + ":" + ExcelUtil.GetCelulaByColumnIndex(3, 9)].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 9) + ":" + ExcelUtil.GetCelulaByColumnIndex(3, 9)].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 9) + ":" + ExcelUtil.GetCelulaByColumnIndex(3, 9)].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 9) + ":" + ExcelUtil.GetCelulaByColumnIndex(3, 9)].Style.Border.Top.Style = ExcelBorderStyle.Thin;

                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 7) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 7) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 7) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 7) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 7) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                        }
                        wworksheet.Select();

                        using (MemoryStream memoryStream = new MemoryStream())
                        {
                            package.SaveAs(memoryStream);

                            memoryStream.Position = 0;

                            returnedMemoryStream = memoryStream;
                        }
                    }
                }

                return File.ReadAllBytes(@"" + fileName);
            }

            catch (Exception)
            {
                throw;
            }
            finally
            {
                File.Delete(fileName);
                File.Delete(templateFileName);
            }
        }

        public byte[] DownloadDashboardTwoEvolutivoMarcasDenominators(GraficoColunasFullLoad grafico, string tituloPlanilha, List<TraducaoComponente> listaTraducaoComponente, FiltroPadraoExcel filtro)
        {
            string fileName = "";
            string templateFileName = "";
            try
            {
                string strDownloadNome;

                strDownloadNome = "Modelo_Exportação_Gráficos_PadraoDenominators";

                string strPath = AdjustPath(HttpContext.Current.Server.MapPath("~")) + @"Arquivos\ExcelExportacao\workspace\";
                fileName = strPath + strDownloadNome + DateTime.Now.ToString("yyyy - MM - dd HH':'mm':'ss':'fff").Replace(" - ", "").Replace(":", "").Replace(" ", "") + ".xlsx";

                templateFileName = AdjustPath(HttpContext.Current.Server.MapPath("~")) + @"Arquivos\ExcelExportacao\Modelos\" + strDownloadNome + ".xlsx";

                GetNewFileMemory(ref templateFileName);

                if (File.Exists(fileName))
                    File.Delete(fileName);

                MemoryStream returnedMemoryStream = null;

                using (var file = new FileStream(fileName, FileMode.CreateNew))
                {
                    using (var temp = new FileStream(templateFileName, FileMode.Open))
                    {
                        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                        ExcelPackage package = new ExcelPackage(file, temp);

                        ExcelWorksheet wworksheet = package.Workbook.Worksheets["Data"];

                        wworksheet.View.ShowHeaders = true;
                        wworksheet.View.ShowGridLines = false;
                        //Titulo da planilha. Deixar linha 3 vazias formatadas no modelo excel.
                        wworksheet.Cells[3, 2].Value = tituloPlanilha;

                        int currentRow = 6;
                        int colNameTable = 3;
                        int lastColName = 0;


                        if (grafico != null)
                        {
                            // Preenche cabeçalho com os periodos:
                            wworksheet.Cells[currentRow, colNameTable].Value = filtro.Onda1.FirstOrDefault().DescItem;
                            colNameTable++;
                            wworksheet.Cells[currentRow, colNameTable].Value = filtro.Onda2.FirstOrDefault().DescItem;
                            colNameTable++;

                            wworksheet.Cells[currentRow, colNameTable].Value = filtro.Onda3.FirstOrDefault().DescItem;
                            colNameTable++;

                            wworksheet.Cells[currentRow, colNameTable].Value = filtro.Onda4.FirstOrDefault().DescItem;
                            colNameTable++;

                            //wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoColunas5.DescMarca;
                            //colNameTable++;

                            colNameTable--;
                            lastColName = colNameTable;

                            //Formata cabeçalho
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(3, 6) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Font.Size = 11;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(3, 6) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Font.Name = "Arial Nova Cond";
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(3, 6) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Font.Bold = true;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(3, 6) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(3, 6) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(3, 6) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(3, 6) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(3, 6) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(3, 6) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                            colNameTable = 2;
                            currentRow++;


                            ////Preenche lateral
                            //wworksheet.Cells[currentRow, colNameTable].Value = "Total Awareness";
                            //currentRow++;
                            //wworksheet.Cells[currentRow, colNameTable].Value = "Prompeted";
                            //currentRow++;
                            //wworksheet.Cells[currentRow, colNameTable].Value = "Spontaneous - OM";
                            //currentRow++;
                            //wworksheet.Cells[currentRow, colNameTable].Value = "Awareness - TOM";
                            //currentRow++;
                            //wworksheet.Cells[currentRow, colNameTable].Value = "Base";

                            wworksheet.Cells[currentRow, colNameTable].Value = listaTraducaoComponente.FirstOrDefault(a => a.Objeto.Equals("dashboard-two-leg-Awareness-total-2")).Texto;
                            currentRow++;
                            wworksheet.Cells[currentRow, colNameTable].Value = listaTraducaoComponente.FirstOrDefault(a => a.Objeto.Equals("dashboard-two-leg-Awareness-total-1")).Texto;
                            currentRow++;


                            wworksheet.Cells[currentRow, colNameTable].Value = "Base";

                            currentRow--;
                            var lastRow = currentRow;
                            colNameTable = 3;
                            currentRow = 7;


                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoColunas1.PercOM;
                            if (grafico.GraficoColunas1.TesteSIGOM == "MENOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Red);
                            if (grafico.GraficoColunas1.TesteSIGOM == "MAIOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Blue);
                            currentRow++;

                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoColunas1.PercTOM;
                            if (grafico.GraficoColunas1.TesteSIGTOM == "MENOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Red);
                            if (grafico.GraficoColunas1.TesteSIGTOM == "MAIOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Blue);
                            currentRow++;

                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoColunas1.BaseAbs;
                            colNameTable++;
                            currentRow = 7;


                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoColunas2.PercOM;
                            if (grafico.GraficoColunas2.TesteSIGOM == "MENOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Red);
                            if (grafico.GraficoColunas2.TesteSIGOM == "MAIOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Blue);
                            currentRow++;

                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoColunas2.PercTOM;
                            if (grafico.GraficoColunas2.TesteSIGTOM == "MENOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Red);
                            if (grafico.GraficoColunas2.TesteSIGTOM == "MAIOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Blue);
                            currentRow++;

                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoColunas2.BaseAbs;
                            colNameTable++;
                            currentRow = 7;


                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoColunas3.PercOM;
                            if (grafico.GraficoColunas3.TesteSIGOM == "MENOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Red);
                            if (grafico.GraficoColunas3.TesteSIGOM == "MAIOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Blue);
                            currentRow++;

                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoColunas3.PercTOM;
                            if (grafico.GraficoColunas3.TesteSIGTOM == "MENOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Red);
                            if (grafico.GraficoColunas3.TesteSIGTOM == "MAIOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Blue);
                            currentRow++;

                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoColunas3.BaseAbs;
                            colNameTable++;
                            currentRow = 7;


                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoColunas4.PercOM;
                            if (grafico.GraficoColunas4.TesteSIGOM == "MENOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Red);
                            if (grafico.GraficoColunas4.TesteSIGOM == "MAIOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Blue);
                            currentRow++;

                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoColunas4.PercTOM;
                            if (grafico.GraficoColunas4.TesteSIGTOM == "MENOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Red);
                            if (grafico.GraficoColunas4.TesteSIGTOM == "MAIOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Blue);
                            currentRow++;

                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoColunas4.BaseAbs;
                            colNameTable++;
                            currentRow = 7;



                            lastRow++;
                            currentRow = lastRow;
                            colNameTable--;
                            lastColName = colNameTable;

                            //Formata dados
                            //wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 7) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Numberformat.Format = "0.00";
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 7) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Font.Size = 11;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 7) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Font.Name = "Arial Nova Cond";
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 7) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Font.Bold = true;

                            // Preenche lateral
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 9) + ":" + ExcelUtil.GetCelulaByColumnIndex(3, 9)].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 9) + ":" + ExcelUtil.GetCelulaByColumnIndex(3, 9)].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 9) + ":" + ExcelUtil.GetCelulaByColumnIndex(3, 9)].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 9) + ":" + ExcelUtil.GetCelulaByColumnIndex(3, 9)].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 9) + ":" + ExcelUtil.GetCelulaByColumnIndex(3, 9)].Style.Border.Top.Style = ExcelBorderStyle.Thin;

                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 7) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 7) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 7) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 7) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 7) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                        }
                        wworksheet.Select();

                        using (MemoryStream memoryStream = new MemoryStream())
                        {
                            package.SaveAs(memoryStream);

                            memoryStream.Position = 0;

                            returnedMemoryStream = memoryStream;
                        }
                    }
                }

                return File.ReadAllBytes(@"" + fileName);
            }

            catch (Exception)
            {
                throw;
            }
            finally
            {
                File.Delete(fileName);
                File.Delete(templateFileName);
            }
        }
        #endregion


        #region DownloadDashboardThree
        public byte[] DownloadDashboardThreeComparativoMarcas(GraficoFunilFullLoad grafico, string tituloPlanilha, List<TraducaoComponente> listaTraducaoComponente)
        {
            string fileName = "";
            string templateFileName = "";
            try
            {
                string strDownloadNome;

                strDownloadNome = "Modelo_Exportação_Gráficos_Padrao_Filtro";

                string strPath = AdjustPath(HttpContext.Current.Server.MapPath("~")) + @"Arquivos\ExcelExportacao\workspace\";
                fileName = strPath + strDownloadNome + DateTime.Now.ToString("yyyy - MM - dd HH':'mm':'ss':'fff").Replace(" - ", "").Replace(":", "").Replace(" ", "") + ".xlsx";

                templateFileName = AdjustPath(HttpContext.Current.Server.MapPath("~")) + @"Arquivos\ExcelExportacao\Modelos\" + strDownloadNome + ".xlsx";

                GetNewFileMemory(ref templateFileName);

                if (File.Exists(fileName))
                    File.Delete(fileName);

                MemoryStream returnedMemoryStream = null;

                using (var file = new FileStream(fileName, FileMode.CreateNew))
                {
                    using (var temp = new FileStream(templateFileName, FileMode.Open))
                    {
                        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                        ExcelPackage package = new ExcelPackage(file, temp);

                        ExcelWorksheet wworksheet = package.Workbook.Worksheets["Data"];

                        wworksheet.View.ShowHeaders = true;
                        wworksheet.View.ShowGridLines = false;
                        //Titulo da planilha. Deixar linha 3 vazias formatadas no modelo excel.
                        wworksheet.Cells[3, 2].Value = tituloPlanilha;

                        int currentRow = 6;
                        int colNameTable = 3;
                        int lastColName = 0;


                        if (grafico != null)
                        {
                            lastColName = colNameTable;

                            colNameTable = 2;
                            currentRow++;

                            ////Preenche lateral

                            wworksheet.Cells[currentRow, colNameTable].Value = listaTraducaoComponente.FirstOrDefault(a => a.Objeto.Equals("grafico-funil-conhecimento")).Texto;
                            currentRow++;
                            wworksheet.Cells[currentRow, colNameTable].Value = listaTraducaoComponente.FirstOrDefault(a => a.Objeto.Equals("grafico-funil-consideracao")).Texto;
                            currentRow++;
                            wworksheet.Cells[currentRow, colNameTable].Value = listaTraducaoComponente.FirstOrDefault(a => a.Objeto.Equals("grafico-funil-uso")).Texto;
                            currentRow++;
                            wworksheet.Cells[currentRow, colNameTable].Value = listaTraducaoComponente.FirstOrDefault(a => a.Objeto.Equals("grafico-funil-preferencia")).Texto;
                            currentRow++;
                            wworksheet.Cells[currentRow, colNameTable].Value = listaTraducaoComponente.FirstOrDefault(a => a.Objeto.Equals("grafico-funil-Loyalty")).Texto;
                            currentRow++;
                            wworksheet.Cells[currentRow, colNameTable].Value = "Base";

                            currentRow--;
                            var lastRow = currentRow;
                            colNameTable = 3;
                            currentRow = 7;

                            #region Coluna 1


                            wworksheet.Cells[currentRow - 2, colNameTable].Value = grafico.GraficoFunil1.DescMarca;
                            wworksheet.Cells[currentRow - 1, colNameTable].Value = grafico.GraficoFunil1.PeriodoAnt;
                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil1.ConhecimentoAnterior;
                            wworksheet.Column(colNameTable).Width = 20;
                            colNameTable++;
                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil1.ConhecimentoAtual;
                            wworksheet.Column(colNameTable).Width = 20;
                            if (grafico.GraficoFunil1.ConhecimentoTesteAtual == "MENOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Red);
                            if (grafico.GraficoFunil1.ConhecimentoTesteAtual == "MAIOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Blue);
                            colNameTable++;

                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil1.ConhecimentoADV;
                            wworksheet.Column(colNameTable).Width = 20;

                            colNameTable++;
                            currentRow++;

                            colNameTable = 3;
                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil1.ConsideracaoAnterior;
                            wworksheet.Column(colNameTable).Width = 20;
                            colNameTable++;
                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil1.ConsideracaoAtual;
                            wworksheet.Column(colNameTable).Width = 20;
                            if (grafico.GraficoFunil1.ConsideracaoTesteAtual == "MENOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Red);
                            if (grafico.GraficoFunil1.ConsideracaoTesteAtual == "MAIOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Blue);
                            colNameTable++;
                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil1.ConsideracaoADV;
                            wworksheet.Column(colNameTable).Width = 20;

                            colNameTable++;
                            currentRow++;

                            colNameTable = 3;
                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil1.UsoAnterior;
                            wworksheet.Column(colNameTable).Width = 20;
                            colNameTable++;
                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil1.UsoAtual;
                            wworksheet.Column(colNameTable).Width = 20;
                            if (grafico.GraficoFunil1.UsoTesteAtual == "MENOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Red);
                            if (grafico.GraficoFunil1.UsoTesteAtual == "MAIOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Blue);
                            colNameTable++;
                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil1.UsoADV;
                            wworksheet.Column(colNameTable).Width = 20;

                            colNameTable++;
                            currentRow++;

                            colNameTable = 3;
                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil1.PreferenciaAnterior;
                            wworksheet.Column(colNameTable).Width = 20;
                            colNameTable++;
                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil1.PreferenciaAtual;
                            wworksheet.Column(colNameTable).Width = 20;
                            if (grafico.GraficoFunil1.PreferenciaTesteAtual == "MENOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Red);
                            if (grafico.GraficoFunil1.PreferenciaTesteAtual == "MAIOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Blue);
                            colNameTable++;
                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil1.PreferenciaADV;
                            wworksheet.Column(colNameTable).Width = 20;

                            colNameTable++;
                            currentRow++;

                            colNameTable = 3;
                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil1.LoyaltyAnterior;
                            wworksheet.Column(colNameTable).Width = 20;
                            colNameTable++;
                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil1.LoyaltyAtual;
                            wworksheet.Column(colNameTable).Width = 20;
                            if (grafico.GraficoFunil1.LoyaltyTesteAtual == "MENOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Red);
                            if (grafico.GraficoFunil1.LoyaltyTesteAtual == "MAIOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Blue);
                            colNameTable++;
                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil1.LoyaltyADV;
                            wworksheet.Column(colNameTable).Width = 20;

                            colNameTable++;
                            currentRow++;

                            colNameTable = 3;
                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil1.BaseAbsAnterior;
                            wworksheet.Column(colNameTable).Width = 20;
                            colNameTable++;
                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil1.BaseAbsAtual;
                            wworksheet.Column(colNameTable).Width = 20;
                            colNameTable++;
                            wworksheet.Cells[currentRow, colNameTable].Value = "";
                            wworksheet.Column(colNameTable).Width = 20;
                            colNameTable++;
                            currentRow++;

                            // Espaço 
                            wworksheet.Cells[currentRow, colNameTable].Value = "";
                            wworksheet.Column(colNameTable).Width = 2;
                            colNameTable++;
                            wworksheet.Cells[currentRow, colNameTable].Value = "";
                            wworksheet.Column(colNameTable).Width = 2;
                            colNameTable++;
                            wworksheet.Cells[currentRow, colNameTable].Value = "";
                            wworksheet.Column(colNameTable).Width = 2;
                            colNameTable++;
                            currentRow++;

                            #endregion

                            #region Coluna 2


                            currentRow = 7;
                            colNameTable = 7;

                            wworksheet.Cells[currentRow - 2, colNameTable].Value = grafico.GraficoFunil2.DescMarca;
                            wworksheet.Cells[currentRow - 1, colNameTable].Value = grafico.GraficoFunil2.PeriodoAnt;
                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil2.ConhecimentoAnterior;
                            wworksheet.Column(colNameTable).Width = 20;
                            colNameTable++;
                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil2.ConhecimentoAtual;
                            wworksheet.Column(colNameTable).Width = 20;
                            if (grafico.GraficoFunil2.ConhecimentoTesteAtual == "MENOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Red);
                            if (grafico.GraficoFunil2.ConhecimentoTesteAtual == "MAIOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Blue);
                            colNameTable++;
                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil2.ConhecimentoADV;
                            wworksheet.Column(colNameTable).Width = 20;

                            colNameTable++;
                            currentRow++;

                            colNameTable = 7;
                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil2.ConsideracaoAnterior;
                            wworksheet.Column(colNameTable).Width = 20;
                            colNameTable++;
                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil2.ConsideracaoAtual;
                            wworksheet.Column(colNameTable).Width = 20;
                            if (grafico.GraficoFunil2.ConsideracaoTesteAtual == "MENOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Red);
                            if (grafico.GraficoFunil2.ConsideracaoTesteAtual == "MAIOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Blue);

                            colNameTable++;

                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil2.ConsideracaoADV;
                            wworksheet.Column(colNameTable).Width = 20;


                            colNameTable++;
                            currentRow++;

                            colNameTable = 7;
                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil2.UsoAnterior;
                            wworksheet.Column(colNameTable).Width = 20;
                            colNameTable++;
                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil2.UsoAtual;
                            wworksheet.Column(colNameTable).Width = 20;
                            if (grafico.GraficoFunil2.UsoTesteAtual == "MENOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Red);
                            if (grafico.GraficoFunil2.UsoTesteAtual == "MAIOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Blue);

                            colNameTable++;

                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil2.UsoADV;
                            wworksheet.Column(colNameTable).Width = 20;
                            colNameTable++;
                            currentRow++;

                            colNameTable = 7;
                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil2.PreferenciaAnterior;
                            wworksheet.Column(colNameTable).Width = 20;
                            colNameTable++;
                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil2.PreferenciaAtual;
                            wworksheet.Column(colNameTable).Width = 20;
                            if (grafico.GraficoFunil2.PreferenciaTesteAtual == "MENOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Red);
                            if (grafico.GraficoFunil2.PreferenciaTesteAtual == "MAIOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Blue);
                            colNameTable++;

                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil2.PreferenciaADV;
                            wworksheet.Column(colNameTable).Width = 20;
                            colNameTable++;
                            currentRow++;

                            colNameTable = 7;
                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil2.LoyaltyAnterior;
                            wworksheet.Column(colNameTable).Width = 20;
                            colNameTable++;
                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil2.LoyaltyAtual;
                            wworksheet.Column(colNameTable).Width = 20;
                            if (grafico.GraficoFunil2.LoyaltyTesteAtual == "MENOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Red);
                            if (grafico.GraficoFunil2.LoyaltyTesteAtual == "MAIOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Blue);
                            colNameTable++;

                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil2.LoyaltyADV;
                            wworksheet.Column(colNameTable).Width = 20;
                            colNameTable++;
                            currentRow++;

                            colNameTable = 7;
                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil2.BaseAbsAnterior;
                            wworksheet.Column(colNameTable).Width = 20;
                            colNameTable++;
                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil2.BaseAbsAtual;
                            wworksheet.Column(colNameTable).Width = 20;
                            colNameTable++;
                            wworksheet.Cells[currentRow, colNameTable].Value = "";
                            wworksheet.Column(colNameTable).Width = 20;
                            colNameTable++;
                            currentRow++;

                            // Espaço 
                            wworksheet.Cells[currentRow, colNameTable].Value = "";
                            wworksheet.Column(colNameTable).Width = 2;
                            colNameTable++;
                            wworksheet.Cells[currentRow, colNameTable].Value = "";
                            wworksheet.Column(colNameTable).Width = 2;
                            colNameTable++;
                            wworksheet.Cells[currentRow, colNameTable].Value = "";
                            wworksheet.Column(colNameTable).Width = 2;
                            colNameTable++;
                            currentRow++;

                            #endregion

                            #region Coluna 3


                            currentRow = 7;
                            colNameTable = 11;

                            wworksheet.Cells[currentRow - 2, colNameTable].Value = grafico.GraficoFunil3.DescMarca;
                            wworksheet.Cells[currentRow - 1, colNameTable].Value = grafico.GraficoFunil3.PeriodoAnt;
                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil3.ConhecimentoAnterior;
                            wworksheet.Column(colNameTable).Width = 20;
                            colNameTable++;
                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil3.ConhecimentoAtual;
                            wworksheet.Column(colNameTable).Width = 20;
                            if (grafico.GraficoFunil3.ConhecimentoTesteAtual == "MENOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Red);
                            if (grafico.GraficoFunil3.ConhecimentoTesteAtual == "MAIOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Blue);
                            colNameTable++;

                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil3.ConhecimentoADV;
                            wworksheet.Column(colNameTable).Width = 20;
                            colNameTable++;
                            currentRow++;

                            colNameTable = 11;
                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil3.ConsideracaoAnterior;
                            wworksheet.Column(colNameTable).Width = 20;
                            colNameTable++;
                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil3.ConsideracaoAtual;
                            wworksheet.Column(colNameTable).Width = 20;
                            if (grafico.GraficoFunil3.ConsideracaoTesteAtual == "MENOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Red);
                            if (grafico.GraficoFunil3.ConsideracaoTesteAtual == "MAIOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Blue);
                            colNameTable++;

                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil3.ConsideracaoADV;
                            wworksheet.Column(colNameTable).Width = 20;
                            colNameTable++;
                            currentRow++;

                            colNameTable = 11;
                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil3.UsoAnterior;
                            wworksheet.Column(colNameTable).Width = 20;
                            colNameTable++;
                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil3.UsoAtual;
                            wworksheet.Column(colNameTable).Width = 20;
                            if (grafico.GraficoFunil3.UsoTesteAtual == "MENOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Red);
                            if (grafico.GraficoFunil3.UsoTesteAtual == "MAIOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Blue);
                            colNameTable++;

                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil3.UsoADV;
                            wworksheet.Column(colNameTable).Width = 20;
                            colNameTable++;
                            currentRow++;

                            colNameTable = 11;
                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil3.PreferenciaAnterior;
                            wworksheet.Column(colNameTable).Width = 20;
                            colNameTable++;
                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil3.PreferenciaAtual;
                            wworksheet.Column(colNameTable).Width = 20;
                            if (grafico.GraficoFunil3.PreferenciaTesteAtual == "MENOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Red);
                            if (grafico.GraficoFunil3.PreferenciaTesteAtual == "MAIOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Blue);
                            colNameTable++;

                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil3.PreferenciaADV;
                            wworksheet.Column(colNameTable).Width = 20;
                            colNameTable++;
                            currentRow++;

                            colNameTable = 11;
                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil3.LoyaltyAnterior;
                            wworksheet.Column(colNameTable).Width = 20;
                            colNameTable++;
                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil3.LoyaltyAtual;
                            wworksheet.Column(colNameTable).Width = 20;
                            if (grafico.GraficoFunil3.LoyaltyTesteAtual == "MENOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Red);
                            if (grafico.GraficoFunil3.LoyaltyTesteAtual == "MAIOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Blue);
                            colNameTable++;

                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil3.LoyaltyADV;
                            wworksheet.Column(colNameTable).Width = 20;
                            colNameTable++;
                            currentRow++;

                            colNameTable = 11;
                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil3.BaseAbsAnterior;
                            wworksheet.Column(colNameTable).Width = 20;
                            colNameTable++;
                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil3.BaseAbsAtual;
                            wworksheet.Column(colNameTable).Width = 20;
                            colNameTable++;
                            wworksheet.Cells[currentRow, colNameTable].Value = "";
                            wworksheet.Column(colNameTable).Width = 20;
                            colNameTable++;
                            currentRow++;

                            // Espaço 
                            wworksheet.Cells[currentRow, colNameTable].Value = "";
                            wworksheet.Column(colNameTable).Width = 2;
                            colNameTable++;
                            wworksheet.Cells[currentRow, colNameTable].Value = "";
                            wworksheet.Column(colNameTable).Width = 2;
                            colNameTable++;
                            wworksheet.Cells[currentRow, colNameTable].Value = "";
                            wworksheet.Column(colNameTable).Width = 2;
                            colNameTable++;
                            currentRow++;

                            #endregion

                            #region Coluna 4


                            currentRow = 7;
                            colNameTable = 15;


                            wworksheet.Cells[currentRow - 2, colNameTable].Value = grafico.GraficoFunil4.DescMarca;
                            wworksheet.Cells[currentRow - 1, colNameTable].Value = grafico.GraficoFunil4.PeriodoAnt;
                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil4.ConhecimentoAnterior;
                            wworksheet.Column(colNameTable).Width = 20;
                            colNameTable++;
                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil4.ConhecimentoAtual;
                            wworksheet.Column(colNameTable).Width = 20;
                            if (grafico.GraficoFunil4.ConhecimentoTesteAtual == "MENOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Red);
                            if (grafico.GraficoFunil4.ConhecimentoTesteAtual == "MAIOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Blue);
                            colNameTable++;

                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil4.ConhecimentoADV;
                            wworksheet.Column(colNameTable).Width = 20;
                            colNameTable++;
                            currentRow++;

                            colNameTable = 15;
                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil4.ConsideracaoAnterior;
                            wworksheet.Column(colNameTable).Width = 20;
                            colNameTable++;
                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil4.ConsideracaoAtual;
                            wworksheet.Column(colNameTable).Width = 20;
                            if (grafico.GraficoFunil4.ConsideracaoTesteAtual == "MENOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Red);
                            if (grafico.GraficoFunil4.ConsideracaoTesteAtual == "MAIOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Blue);
                            colNameTable++;

                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil4.ConsideracaoADV;
                            wworksheet.Column(colNameTable).Width = 20;
                            colNameTable++;
                            currentRow++;

                            colNameTable = 15;
                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil4.UsoAnterior;
                            wworksheet.Column(colNameTable).Width = 20;
                            colNameTable++;
                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil4.UsoAtual;
                            wworksheet.Column(colNameTable).Width = 20;
                            if (grafico.GraficoFunil4.UsoTesteAtual == "MENOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Red);
                            if (grafico.GraficoFunil4.UsoTesteAtual == "MAIOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Blue);
                            colNameTable++;

                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil4.UsoADV;
                            wworksheet.Column(colNameTable).Width = 20;
                            colNameTable++;
                            currentRow++;

                            colNameTable = 15;
                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil4.PreferenciaAnterior;
                            wworksheet.Column(colNameTable).Width = 20;
                            colNameTable++;
                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil4.PreferenciaAtual;
                            wworksheet.Column(colNameTable).Width = 20;
                            if (grafico.GraficoFunil4.PreferenciaTesteAtual == "MENOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Red);
                            if (grafico.GraficoFunil4.PreferenciaTesteAtual == "MAIOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Blue);
                            colNameTable++;

                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil4.PreferenciaADV;
                            wworksheet.Column(colNameTable).Width = 20;
                            colNameTable++;
                            currentRow++;


                            colNameTable = 15;
                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil4.LoyaltyAnterior;
                            wworksheet.Column(colNameTable).Width = 20;
                            colNameTable++;
                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil4.LoyaltyAtual;
                            wworksheet.Column(colNameTable).Width = 20;
                            if (grafico.GraficoFunil4.LoyaltyTesteAtual == "MENOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Red);
                            if (grafico.GraficoFunil4.LoyaltyTesteAtual == "MAIOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Blue);
                            colNameTable++;

                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil4.LoyaltyADV;
                            wworksheet.Column(colNameTable).Width = 20;
                            colNameTable++;
                            currentRow++;


                            colNameTable = 15;
                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil4.BaseAbsAnterior;
                            wworksheet.Column(colNameTable).Width = 20;
                            colNameTable++;
                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil4.BaseAbsAtual;
                            wworksheet.Column(colNameTable).Width = 20;
                            colNameTable++;
                            wworksheet.Cells[currentRow, colNameTable].Value = "";
                            wworksheet.Column(colNameTable).Width = 20;
                            colNameTable++;
                            currentRow++;

                            // Espaço 
                            wworksheet.Cells[currentRow, colNameTable].Value = "";
                            wworksheet.Column(colNameTable).Width = 2;
                            colNameTable++;
                            wworksheet.Cells[currentRow, colNameTable].Value = "";
                            wworksheet.Column(colNameTable).Width = 2;
                            colNameTable++;
                            wworksheet.Cells[currentRow, colNameTable].Value = "";
                            wworksheet.Column(colNameTable).Width = 2;
                            colNameTable++;
                            currentRow++;

                            #endregion


                            colNameTable = 2;
                            currentRow = 15;

                            ////Preenche lateral

                            wworksheet.Cells[currentRow, colNameTable].Value = listaTraducaoComponente.FirstOrDefault(a => a.Objeto.Equals("grafico-funil-conhecimento")).Texto;
                            currentRow++;
                            wworksheet.Cells[currentRow, colNameTable].Value = listaTraducaoComponente.FirstOrDefault(a => a.Objeto.Equals("grafico-funil-consideracao")).Texto;
                            currentRow++;
                            wworksheet.Cells[currentRow, colNameTable].Value = listaTraducaoComponente.FirstOrDefault(a => a.Objeto.Equals("grafico-funil-uso")).Texto;
                            currentRow++;
                            wworksheet.Cells[currentRow, colNameTable].Value = listaTraducaoComponente.FirstOrDefault(a => a.Objeto.Equals("grafico-funil-preferencia")).Texto;
                            currentRow++;
                            wworksheet.Cells[currentRow, colNameTable].Value = listaTraducaoComponente.FirstOrDefault(a => a.Objeto.Equals("grafico-funil-Loyalty")).Texto;
                            currentRow++;
                            wworksheet.Cells[currentRow, colNameTable].Value = "Base";


                            #region Coluna 5


                            currentRow = 15;
                            colNameTable = 3;
                            wworksheet.Cells[currentRow - 2, colNameTable].Value = grafico.GraficoFunil5.DescMarca;
                            wworksheet.Cells[currentRow - 1, colNameTable].Value = grafico.GraficoFunil5.PeriodoAnt;
                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil5.ConhecimentoAnterior;
                            wworksheet.Column(colNameTable).Width = 20;
                            colNameTable++;
                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil5.ConhecimentoAtual;
                            wworksheet.Column(colNameTable).Width = 20;
                            if (grafico.GraficoFunil5.ConhecimentoTesteAtual == "MENOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Red);
                            if (grafico.GraficoFunil5.ConhecimentoTesteAtual == "MAIOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Blue);
                            colNameTable++;

                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil5.ConhecimentoADV;
                            wworksheet.Column(colNameTable).Width = 20;
                            colNameTable++;
                            currentRow++;

                            colNameTable = 3;
                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil5.ConsideracaoAnterior;
                            wworksheet.Column(colNameTable).Width = 20;
                            colNameTable++;
                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil5.ConsideracaoAtual;
                            wworksheet.Column(colNameTable).Width = 20;
                            if (grafico.GraficoFunil5.ConsideracaoTesteAtual == "MENOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Red);
                            if (grafico.GraficoFunil5.ConsideracaoTesteAtual == "MAIOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Blue);
                            colNameTable++;

                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil5.ConsideracaoADV;
                            wworksheet.Column(colNameTable).Width = 20;
                            colNameTable++;
                            currentRow++;

                            colNameTable = 3;
                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil5.UsoAnterior;
                            wworksheet.Column(colNameTable).Width = 20;
                            colNameTable++;
                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil5.UsoAtual;
                            wworksheet.Column(colNameTable).Width = 20;
                            if (grafico.GraficoFunil5.UsoTesteAtual == "MENOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Red);
                            if (grafico.GraficoFunil5.UsoTesteAtual == "MAIOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Blue);
                            colNameTable++;

                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil5.UsoADV;
                            wworksheet.Column(colNameTable).Width = 20;
                            colNameTable++;
                            currentRow++;

                            colNameTable = 3;
                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil5.PreferenciaAnterior;
                            wworksheet.Column(colNameTable).Width = 20;
                            colNameTable++;
                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil5.PreferenciaAtual;
                            wworksheet.Column(colNameTable).Width = 20;
                            if (grafico.GraficoFunil5.PreferenciaTesteAtual == "MENOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Red);
                            if (grafico.GraficoFunil5.PreferenciaTesteAtual == "MAIOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Blue);
                            colNameTable++;

                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil5.PreferenciaADV;
                            wworksheet.Column(colNameTable).Width = 20;
                            colNameTable++;
                            currentRow++;

                            colNameTable = 3;
                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil5.LoyaltyAnterior;
                            wworksheet.Column(colNameTable).Width = 20;
                            colNameTable++;
                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil5.LoyaltyAtual;
                            wworksheet.Column(colNameTable).Width = 20;
                            if (grafico.GraficoFunil5.LoyaltyTesteAtual == "MENOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Red);
                            if (grafico.GraficoFunil5.LoyaltyTesteAtual == "MAIOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Blue);
                            colNameTable++;

                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil5.LoyaltyADV;
                            wworksheet.Column(colNameTable).Width = 20;
                            colNameTable++;
                            currentRow++;

                            colNameTable = 3;
                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil5.BaseAbsAnterior;
                            wworksheet.Column(colNameTable).Width = 20;
                            colNameTable++;
                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil5.BaseAbsAtual;
                            wworksheet.Column(colNameTable).Width = 20;
                            colNameTable++;
                            wworksheet.Cells[currentRow, colNameTable].Value = "";
                            wworksheet.Column(colNameTable).Width = 20;

                            colNameTable++;
                            currentRow++;

                            // Espaço 
                            wworksheet.Cells[currentRow, colNameTable].Value = "";
                            wworksheet.Column(colNameTable).Width = 2;
                            colNameTable++;
                            wworksheet.Cells[currentRow, colNameTable].Value = "";
                            wworksheet.Column(colNameTable).Width = 2;
                            colNameTable++;
                            wworksheet.Cells[currentRow, colNameTable].Value = "";
                            wworksheet.Column(colNameTable).Width = 2;
                            colNameTable++;
                            currentRow++;

                            #endregion

                            #region Coluna 6


                            currentRow = 15;
                            colNameTable = 7;
                            wworksheet.Cells[currentRow - 2, colNameTable].Value = grafico.GraficoFunil6.DescMarca;
                            wworksheet.Cells[currentRow - 1, colNameTable].Value = grafico.GraficoFunil6.PeriodoAnt;
                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil6.ConhecimentoAnterior;
                            wworksheet.Column(colNameTable).Width = 20;
                            colNameTable++;
                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil6.ConhecimentoAtual;
                            wworksheet.Column(colNameTable).Width = 20;
                            if (grafico.GraficoFunil6.ConhecimentoTesteAtual == "MENOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Red);
                            if (grafico.GraficoFunil6.ConhecimentoTesteAtual == "MAIOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Blue);
                            colNameTable++;

                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil6.ConhecimentoADV;
                            wworksheet.Column(colNameTable).Width = 20;
                            colNameTable++;
                            currentRow++;

                            colNameTable = 7;
                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil6.ConsideracaoAnterior;
                            wworksheet.Column(colNameTable).Width = 20;
                            colNameTable++;
                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil6.ConsideracaoAtual;
                            wworksheet.Column(colNameTable).Width = 20;
                            if (grafico.GraficoFunil6.ConsideracaoTesteAtual == "MENOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Red);
                            if (grafico.GraficoFunil6.ConsideracaoTesteAtual == "MAIOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Blue);
                            colNameTable++;

                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil6.ConsideracaoADV;
                            wworksheet.Column(colNameTable).Width = 20;
                            colNameTable++;
                            currentRow++;

                            colNameTable = 7;
                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil6.UsoAnterior;
                            wworksheet.Column(colNameTable).Width = 20;
                            colNameTable++;
                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil6.UsoAtual;
                            wworksheet.Column(colNameTable).Width = 20;
                            if (grafico.GraficoFunil6.UsoTesteAtual == "MENOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Red);
                            if (grafico.GraficoFunil6.UsoTesteAtual == "MAIOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Blue);
                            colNameTable++;

                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil6.UsoADV;
                            wworksheet.Column(colNameTable).Width = 20;
                            colNameTable++;
                            currentRow++;

                            colNameTable = 7;
                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil6.PreferenciaAnterior;
                            wworksheet.Column(colNameTable).Width = 20;
                            colNameTable++;
                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil6.PreferenciaAtual;
                            wworksheet.Column(colNameTable).Width = 20;
                            if (grafico.GraficoFunil6.PreferenciaTesteAtual == "MENOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Red);
                            if (grafico.GraficoFunil6.PreferenciaTesteAtual == "MAIOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Blue);
                            colNameTable++;

                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil6.PreferenciaADV;
                            wworksheet.Column(colNameTable).Width = 20;
                            colNameTable++;
                            currentRow++;


                            colNameTable = 7;
                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil6.LoyaltyAnterior;
                            wworksheet.Column(colNameTable).Width = 20;
                            colNameTable++;
                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil6.LoyaltyAtual;
                            wworksheet.Column(colNameTable).Width = 20;
                            if (grafico.GraficoFunil6.LoyaltyTesteAtual == "MENOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Red);
                            if (grafico.GraficoFunil6.LoyaltyTesteAtual == "MAIOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Blue);
                            colNameTable++;

                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil6.LoyaltyADV;
                            wworksheet.Column(colNameTable).Width = 20;
                            colNameTable++;
                            currentRow++;

                            colNameTable = 7;
                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil6.BaseAbsAnterior;
                            wworksheet.Column(colNameTable).Width = 20;
                            colNameTable++;
                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil6.BaseAbsAtual;
                            wworksheet.Column(colNameTable).Width = 20;
                            colNameTable++;
                            wworksheet.Cells[currentRow, colNameTable].Value = "";
                            wworksheet.Column(colNameTable).Width = 20;
                            colNameTable++;
                            currentRow++;

                            // Espaço 
                            wworksheet.Cells[currentRow, colNameTable].Value = "";
                            wworksheet.Column(colNameTable).Width = 2;
                            colNameTable++;
                            wworksheet.Cells[currentRow, colNameTable].Value = "";
                            wworksheet.Column(colNameTable).Width = 2;
                            colNameTable++;
                            wworksheet.Cells[currentRow, colNameTable].Value = "";
                            wworksheet.Column(colNameTable).Width = 2;
                            colNameTable++;
                            currentRow++;

                            #endregion

                            #region Coluna 7


                            currentRow = 15;
                            colNameTable = 11;
                            wworksheet.Cells[currentRow - 2, colNameTable].Value = grafico.GraficoFunil7.DescMarca;
                            wworksheet.Cells[currentRow - 1, colNameTable].Value = grafico.GraficoFunil7.PeriodoAnt;
                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil7.ConhecimentoAnterior;
                            wworksheet.Column(colNameTable).Width = 20;
                            colNameTable++;
                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil7.ConhecimentoAtual;
                            wworksheet.Column(colNameTable).Width = 20;
                            if (grafico.GraficoFunil7.ConhecimentoTesteAtual == "MENOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Red);
                            if (grafico.GraficoFunil7.ConhecimentoTesteAtual == "MAIOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Blue);
                            colNameTable++;

                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil7.ConhecimentoADV;
                            wworksheet.Column(colNameTable).Width = 20;
                            colNameTable++;
                            currentRow++;

                            colNameTable = 11;
                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil7.ConsideracaoAnterior;
                            wworksheet.Column(colNameTable).Width = 20;
                            colNameTable++;
                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil7.ConsideracaoAtual;
                            wworksheet.Column(colNameTable).Width = 20;
                            if (grafico.GraficoFunil7.ConsideracaoTesteAtual == "MENOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Red);
                            if (grafico.GraficoFunil7.ConsideracaoTesteAtual == "MAIOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Blue);
                            colNameTable++;

                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil7.ConsideracaoADV;
                            wworksheet.Column(colNameTable).Width = 20;
                            colNameTable++;
                            currentRow++;

                            colNameTable = 11;
                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil7.UsoAnterior;
                            wworksheet.Column(colNameTable).Width = 20;
                            colNameTable++;
                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil7.UsoAtual;
                            wworksheet.Column(colNameTable).Width = 20;
                            if (grafico.GraficoFunil7.UsoTesteAtual == "MENOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Red);
                            if (grafico.GraficoFunil7.UsoTesteAtual == "MAIOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Blue);
                            colNameTable++;

                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil7.UsoADV;
                            wworksheet.Column(colNameTable).Width = 20;
                            colNameTable++;
                            currentRow++;

                            colNameTable = 11;
                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil7.PreferenciaAnterior;
                            wworksheet.Column(colNameTable).Width = 20;
                            colNameTable++;
                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil7.PreferenciaAtual;
                            wworksheet.Column(colNameTable).Width = 20;
                            if (grafico.GraficoFunil7.PreferenciaTesteAtual == "MENOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Red);
                            if (grafico.GraficoFunil7.PreferenciaTesteAtual == "MAIOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Blue);
                            colNameTable++;

                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil7.PreferenciaADV;
                            wworksheet.Column(colNameTable).Width = 20;
                            colNameTable++;
                            currentRow++;

                            colNameTable = 11;
                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil7.LoyaltyAnterior;
                            wworksheet.Column(colNameTable).Width = 20;
                            colNameTable++;
                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil7.LoyaltyAtual;
                            wworksheet.Column(colNameTable).Width = 20;
                            if (grafico.GraficoFunil7.LoyaltyTesteAtual == "MENOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Red);
                            if (grafico.GraficoFunil7.LoyaltyTesteAtual == "MAIOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Blue);
                            colNameTable++;

                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil7.LoyaltyADV;
                            wworksheet.Column(colNameTable).Width = 20;
                            colNameTable++;
                            currentRow++;

                            colNameTable = 11;
                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil7.BaseAbsAnterior;
                            wworksheet.Column(colNameTable).Width = 20;
                            colNameTable++;
                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil7.BaseAbsAtual;
                            wworksheet.Column(colNameTable).Width = 20;
                            colNameTable++;
                            wworksheet.Cells[currentRow, colNameTable].Value = "";
                            wworksheet.Column(colNameTable).Width = 20;
                            colNameTable++;
                            currentRow++;

                            // Espaço 
                            wworksheet.Cells[currentRow, colNameTable].Value = "";
                            wworksheet.Column(colNameTable).Width = 2;
                            colNameTable++;
                            wworksheet.Cells[currentRow, colNameTable].Value = "";
                            wworksheet.Column(colNameTable).Width = 2;
                            colNameTable++;
                            wworksheet.Cells[currentRow, colNameTable].Value = "";
                            wworksheet.Column(colNameTable).Width = 2;
                            colNameTable++;
                            currentRow++;

                            #endregion

                            #region Coluna 8


                            currentRow = 15;
                            colNameTable = 15;
                            wworksheet.Cells[currentRow - 2, colNameTable].Value = grafico.GraficoFunil8.DescMarca;
                            wworksheet.Cells[currentRow - 1, colNameTable].Value = grafico.GraficoFunil8.PeriodoAnt;
                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil8.ConhecimentoAnterior;
                            wworksheet.Column(colNameTable).Width = 20;
                            colNameTable++;
                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil8.ConhecimentoAtual;
                            wworksheet.Column(colNameTable).Width = 20;
                            if (grafico.GraficoFunil8.ConhecimentoTesteAtual == "MENOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Red);
                            if (grafico.GraficoFunil8.ConhecimentoTesteAtual == "MAIOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Blue);
                            colNameTable++;

                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil8.ConhecimentoADV;
                            wworksheet.Column(colNameTable).Width = 20;
                            colNameTable++;
                            currentRow++;

                            colNameTable = 15;
                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil8.ConsideracaoAnterior;
                            wworksheet.Column(colNameTable).Width = 20;
                            colNameTable++;
                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil8.ConsideracaoAtual;
                            wworksheet.Column(colNameTable).Width = 20;
                            if (grafico.GraficoFunil8.ConsideracaoTesteAtual == "MENOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Red);
                            if (grafico.GraficoFunil8.ConsideracaoTesteAtual == "MAIOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Blue);
                            colNameTable++;

                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil8.ConsideracaoADV;
                            wworksheet.Column(colNameTable).Width = 20;
                            colNameTable++;
                            currentRow++;

                            colNameTable = 15;
                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil8.UsoAnterior;
                            wworksheet.Column(colNameTable).Width = 20;
                            colNameTable++;
                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil8.UsoAtual;
                            wworksheet.Column(colNameTable).Width = 20;
                            if (grafico.GraficoFunil8.UsoTesteAtual == "MENOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Red);
                            if (grafico.GraficoFunil8.UsoTesteAtual == "MAIOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Blue);
                            colNameTable++;

                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil8.UsoADV;
                            wworksheet.Column(colNameTable).Width = 20;
                            colNameTable++;
                            currentRow++;

                            colNameTable = 15;
                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil8.PreferenciaAnterior;
                            wworksheet.Column(colNameTable).Width = 20;
                            colNameTable++;
                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil8.PreferenciaAtual;
                            wworksheet.Column(colNameTable).Width = 20;
                            if (grafico.GraficoFunil8.PreferenciaTesteAtual == "MENOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Red);
                            if (grafico.GraficoFunil8.PreferenciaTesteAtual == "MAIOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Blue);
                            colNameTable++;

                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil8.PreferenciaADV;
                            wworksheet.Column(colNameTable).Width = 20;
                            colNameTable++;
                            currentRow++;

                            colNameTable = 15;
                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil8.LoyaltyAnterior;
                            wworksheet.Column(colNameTable).Width = 20;
                            colNameTable++;
                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil8.LoyaltyAtual;
                            wworksheet.Column(colNameTable).Width = 20;
                            if (grafico.GraficoFunil8.LoyaltyTesteAtual == "MENOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Red);
                            if (grafico.GraficoFunil8.LoyaltyTesteAtual == "MAIOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Blue);
                            colNameTable++;

                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil8.LoyaltyADV;
                            wworksheet.Column(colNameTable).Width = 20;
                            colNameTable++;
                            currentRow++;

                            colNameTable = 15;
                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil8.BaseAbsAnterior;
                            wworksheet.Column(colNameTable).Width = 20;
                            colNameTable++;
                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil8.BaseAbsAtual;
                            wworksheet.Column(colNameTable).Width = 20;
                            colNameTable++;
                            wworksheet.Cells[currentRow, colNameTable].Value = "";
                            wworksheet.Column(colNameTable).Width = 20;

                            colNameTable++;
                            currentRow++;

                            // Espaço 
                            wworksheet.Cells[currentRow, colNameTable].Value = "";
                            wworksheet.Column(colNameTable).Width = 2;
                            colNameTable++;
                            wworksheet.Cells[currentRow, colNameTable].Value = "";
                            wworksheet.Column(colNameTable).Width = 2;
                            colNameTable++;
                            wworksheet.Cells[currentRow, colNameTable].Value = "";
                            wworksheet.Column(colNameTable).Width = 2;
                            colNameTable++;
                            currentRow++;

                            #endregion

                            currentRow = 7;

                            lastRow++;
                            currentRow = lastRow;
                            colNameTable--;
                            lastColName = colNameTable;

                            //Formata dados
                            //wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 7) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Numberformat.Format = "0.00";
                            //wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 7) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Numberformat.Format = "0.0";
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 7) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Font.Size = 11;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 7) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Font.Name = "Arial Nova Cond";
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 7) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Font.Bold = true;


                        }
                        wworksheet.Select();

                        using (MemoryStream memoryStream = new MemoryStream())
                        {
                            package.SaveAs(memoryStream);

                            memoryStream.Position = 0;

                            returnedMemoryStream = memoryStream;
                        }
                    }
                }

                return File.ReadAllBytes(@"" + fileName);
            }

            catch (Exception)
            {
                throw;
            }
            finally
            {
                File.Delete(fileName);
                File.Delete(templateFileName);
            }
        }

        public byte[] DownloadDashboardThreeEvolutivoMarcas(GraficoFunilFullLoad grafico, string tituloPlanilha, List<TraducaoComponente> listaTraducaoComponente, FiltroPadraoExcel filtro)
        {
            string fileName = "";
            string templateFileName = "";
            try
            {
                string strDownloadNome;

                strDownloadNome = "Modelo_Exportação_Gráficos_Padrao_Filtro_Unico";

                string strPath = AdjustPath(HttpContext.Current.Server.MapPath("~")) + @"Arquivos\ExcelExportacao\workspace\";
                fileName = strPath + strDownloadNome + DateTime.Now.ToString("yyyy - MM - dd HH':'mm':'ss':'fff").Replace(" - ", "").Replace(":", "").Replace(" ", "") + ".xlsx";

                templateFileName = AdjustPath(HttpContext.Current.Server.MapPath("~")) + @"Arquivos\ExcelExportacao\Modelos\" + strDownloadNome + ".xlsx";

                GetNewFileMemory(ref templateFileName);

                if (File.Exists(fileName))
                    File.Delete(fileName);

                MemoryStream returnedMemoryStream = null;

                using (var file = new FileStream(fileName, FileMode.CreateNew))
                {
                    using (var temp = new FileStream(templateFileName, FileMode.Open))
                    {
                        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                        ExcelPackage package = new ExcelPackage(file, temp);

                        ExcelWorksheet wworksheet = package.Workbook.Worksheets["Data"];

                        wworksheet.View.ShowHeaders = true;
                        wworksheet.View.ShowGridLines = false;
                        //Titulo da planilha. Deixar linha 3 vazias formatadas no modelo excel.
                        wworksheet.Cells[3, 2].Value = tituloPlanilha;

                        int currentRow = 5;
                        int colNameTable = 3;
                        int lastColName = 0;


                        if (grafico != null)
                        {
                            lastColName = colNameTable;

                            colNameTable = 2;


                            ////Preenche lateral
                            ///

                            if (filtro.Marca != null && filtro.Marca.Any())
                            {
                                Color colour = ColorTranslator.FromHtml(filtro.Marca?.FirstOrDefault().CorItem);
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(colour);
                                wworksheet.Cells[currentRow, colNameTable].Value = filtro.Marca?.FirstOrDefault().DescItem;
                                currentRow++;
                            }
                            else
                            {

                                wworksheet.Cells[currentRow, colNameTable].Value = "";
                                currentRow++;
                            }


                            wworksheet.Cells[currentRow, colNameTable].Value = listaTraducaoComponente.FirstOrDefault(a => a.Objeto.Equals("grafico-funil-conhecimento")).Texto;
                            currentRow++;
                            wworksheet.Cells[currentRow, colNameTable].Value = listaTraducaoComponente.FirstOrDefault(a => a.Objeto.Equals("grafico-funil-consideracao")).Texto;
                            currentRow++;
                            wworksheet.Cells[currentRow, colNameTable].Value = listaTraducaoComponente.FirstOrDefault(a => a.Objeto.Equals("grafico-funil-uso")).Texto;
                            currentRow++;
                            wworksheet.Cells[currentRow, colNameTable].Value = listaTraducaoComponente.FirstOrDefault(a => a.Objeto.Equals("grafico-funil-preferencia")).Texto;
                            currentRow++;
                            wworksheet.Cells[currentRow, colNameTable].Value = listaTraducaoComponente.FirstOrDefault(a => a.Objeto.Equals("grafico-funil-Loyalty")).Texto;
                            currentRow++;
                            wworksheet.Cells[currentRow, colNameTable].Value = "Base";

                            currentRow--;
                            var lastRow = currentRow;
                            colNameTable = 3;
                            currentRow = 6;

                            #region Coluna 1

                            if (filtro.Onda1.FirstOrDefault() == null)
                                wworksheet.Cells[currentRow - 1, colNameTable].Value = "N/A";
                            else
                                wworksheet.Cells[currentRow - 1, colNameTable].Value = filtro.Onda1.FirstOrDefault().DescItem;

                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil1.ConhecimentoAtual;
                            wworksheet.Column(colNameTable).Width = 20;
                            if (grafico.GraficoFunil1.ConhecimentoTesteAtual == "MENOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Red);
                            if (grafico.GraficoFunil1.ConhecimentoTesteAtual == "MAIOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Blue);
                            colNameTable++;

                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil1.ConhecimentoADV;
                            wworksheet.Column(colNameTable).Width = 20;
                            colNameTable++;
                            currentRow++;

                            colNameTable = 3;

                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil1.ConsideracaoAtual;
                            wworksheet.Column(colNameTable).Width = 20;
                            if (grafico.GraficoFunil1.ConsideracaoTesteAtual == "MENOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Red);
                            if (grafico.GraficoFunil1.ConsideracaoTesteAtual == "MAIOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Blue);
                            colNameTable++;

                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil1.ConsideracaoADV;
                            wworksheet.Column(colNameTable).Width = 20;
                            colNameTable++;
                            currentRow++;

                            colNameTable = 3;

                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil1.UsoAtual;
                            wworksheet.Column(colNameTable).Width = 20;
                            if (grafico.GraficoFunil1.UsoTesteAtual == "MENOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Red);
                            if (grafico.GraficoFunil1.UsoTesteAtual == "MAIOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Blue);
                            colNameTable++;

                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil1.UsoADV;
                            wworksheet.Column(colNameTable).Width = 20;
                            colNameTable++;
                            currentRow++;

                            colNameTable = 3;

                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil1.PreferenciaAtual;
                            wworksheet.Column(colNameTable).Width = 20;
                            if (grafico.GraficoFunil1.PreferenciaTesteAtual == "MENOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Red);
                            if (grafico.GraficoFunil1.PreferenciaTesteAtual == "MAIOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Blue);
                            colNameTable++;

                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil1.PreferenciaADV;
                            wworksheet.Column(colNameTable).Width = 20;
                            colNameTable++;
                            currentRow++;

                            colNameTable = 3;

                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil1.LoyaltyAtual;
                            wworksheet.Column(colNameTable).Width = 20;
                            if (grafico.GraficoFunil1.LoyaltyTesteAtual == "MENOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Red);
                            if (grafico.GraficoFunil1.LoyaltyTesteAtual == "MAIOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Blue);
                            colNameTable++;

                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil1.LoyaltyADV;
                            wworksheet.Column(colNameTable).Width = 20;
                            colNameTable++;
                            currentRow++;

                            colNameTable = 3;

                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil1.BaseAbsAtual;
                            wworksheet.Column(colNameTable).Width = 20;
                            colNameTable++;
                            wworksheet.Cells[currentRow, colNameTable].Value = "";
                            wworksheet.Column(colNameTable).Width = 20;
                            colNameTable++;
                            currentRow++;

                            // Espaço 
                            wworksheet.Cells[currentRow, colNameTable].Value = "";
                            wworksheet.Column(colNameTable).Width = 2;
                            colNameTable++;
                            wworksheet.Cells[currentRow, colNameTable].Value = "";
                            wworksheet.Column(colNameTable).Width = 2;
                            colNameTable++;
                            wworksheet.Cells[currentRow, colNameTable].Value = "";
                            wworksheet.Column(colNameTable).Width = 2;
                            colNameTable++;
                            currentRow++;

                            #endregion

                            #region Coluna 2


                            currentRow = 6;
                            colNameTable = 6;

                            if (filtro.Onda2.FirstOrDefault() == null)
                                wworksheet.Cells[currentRow - 1, colNameTable].Value = "N/A";
                            else
                                wworksheet.Cells[currentRow - 1, colNameTable].Value = filtro.Onda2.FirstOrDefault().DescItem;

                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil2.ConhecimentoAtual;
                            wworksheet.Column(colNameTable).Width = 20;
                            if (grafico.GraficoFunil2.ConhecimentoTesteAtual == "MENOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Red);
                            if (grafico.GraficoFunil2.ConhecimentoTesteAtual == "MAIOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Blue);
                            colNameTable++;

                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil2.ConhecimentoADV;
                            wworksheet.Column(colNameTable).Width = 20;
                            colNameTable++;
                            currentRow++;

                            colNameTable = 6;

                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil2.ConsideracaoAtual;
                            wworksheet.Column(colNameTable).Width = 20;
                            if (grafico.GraficoFunil2.ConsideracaoTesteAtual == "MENOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Red);
                            if (grafico.GraficoFunil2.ConsideracaoTesteAtual == "MAIOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Blue);
                            colNameTable++;

                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil2.ConsideracaoADV;
                            wworksheet.Column(colNameTable).Width = 20;
                            colNameTable++;
                            currentRow++;

                            colNameTable = 6;

                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil2.UsoAtual;
                            wworksheet.Column(colNameTable).Width = 20;
                            if (grafico.GraficoFunil2.UsoTesteAtual == "MENOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Red);
                            if (grafico.GraficoFunil2.UsoTesteAtual == "MAIOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Blue);
                            colNameTable++;

                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil2.UsoADV;
                            wworksheet.Column(colNameTable).Width = 20;
                            colNameTable++;
                            currentRow++;

                            colNameTable = 6;

                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil2.PreferenciaAtual;
                            wworksheet.Column(colNameTable).Width = 20;
                            if (grafico.GraficoFunil2.PreferenciaTesteAtual == "MENOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Red);
                            if (grafico.GraficoFunil2.PreferenciaTesteAtual == "MAIOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Blue);
                            colNameTable++;

                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil2.PreferenciaADV;
                            wworksheet.Column(colNameTable).Width = 20;
                            colNameTable++;
                            currentRow++;

                            colNameTable = 6;

                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil2.LoyaltyAtual;
                            wworksheet.Column(colNameTable).Width = 20;
                            if (grafico.GraficoFunil2.LoyaltyTesteAtual == "MENOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Red);
                            if (grafico.GraficoFunil2.LoyaltyTesteAtual == "MAIOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Blue);
                            colNameTable++;

                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil2.LoyaltyADV;
                            wworksheet.Column(colNameTable).Width = 20;
                            colNameTable++;
                            currentRow++;

                            colNameTable = 6;

                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil2.BaseAbsAtual;
                            wworksheet.Column(colNameTable).Width = 20;
                            colNameTable++;
                            wworksheet.Cells[currentRow, colNameTable].Value = "";
                            wworksheet.Column(colNameTable).Width = 20;
                            colNameTable++;
                            currentRow++;

                            // Espaço 
                            wworksheet.Cells[currentRow, colNameTable].Value = "";
                            wworksheet.Column(colNameTable).Width = 2;
                            colNameTable++;
                            wworksheet.Cells[currentRow, colNameTable].Value = "";
                            wworksheet.Column(colNameTable).Width = 2;
                            colNameTable++;
                            wworksheet.Cells[currentRow, colNameTable].Value = "";
                            wworksheet.Column(colNameTable).Width = 2;
                            colNameTable++;
                            currentRow++;

                            #endregion

                            #region Coluna 3


                            currentRow = 6;
                            colNameTable = 9;

                   
                            if (filtro.Onda3.FirstOrDefault() == null)
                                wworksheet.Cells[currentRow - 1, colNameTable].Value = "N/A";
                            else
                                wworksheet.Cells[currentRow - 1, colNameTable].Value = filtro.Onda3.FirstOrDefault().DescItem;

                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil3.ConhecimentoAtual;
                            wworksheet.Column(colNameTable).Width = 20;
                            if (grafico.GraficoFunil3.ConhecimentoTesteAtual == "MENOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Red);
                            if (grafico.GraficoFunil3.ConhecimentoTesteAtual == "MAIOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Blue);
                            colNameTable++;

                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil3.ConhecimentoADV;
                            wworksheet.Column(colNameTable).Width = 20;
                            colNameTable++;
                            currentRow++;

                            colNameTable = 9;

                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil3.ConsideracaoAtual;
                            wworksheet.Column(colNameTable).Width = 20;

                            if (grafico.GraficoFunil3.ConsideracaoTesteAtual == "MENOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Red);
                            if (grafico.GraficoFunil3.ConsideracaoTesteAtual == "MAIOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Blue);
                            colNameTable++;

                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil3.ConsideracaoADV;
                            wworksheet.Column(colNameTable).Width = 20;
                            colNameTable++;
                            currentRow++;

                            colNameTable = 9;

                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil3.UsoAtual;
                            wworksheet.Column(colNameTable).Width = 20;
                            if (grafico.GraficoFunil3.UsoTesteAtual == "MENOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Red);
                            if (grafico.GraficoFunil3.UsoTesteAtual == "MAIOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Blue);
                            colNameTable++;

                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil3.UsoADV;
                            wworksheet.Column(colNameTable).Width = 20;
                            colNameTable++;
                            currentRow++;

                            colNameTable = 9;

                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil3.PreferenciaAtual;
                            wworksheet.Column(colNameTable).Width = 20;
                            if (grafico.GraficoFunil3.PreferenciaTesteAtual == "MENOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Red);
                            if (grafico.GraficoFunil3.PreferenciaTesteAtual == "MAIOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Blue);
                            colNameTable++;

                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil3.PreferenciaADV;
                            wworksheet.Column(colNameTable).Width = 20;
                            colNameTable++;
                            currentRow++;

                            colNameTable = 9;

                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil3.LoyaltyAtual;
                            wworksheet.Column(colNameTable).Width = 20;
                            if (grafico.GraficoFunil3.LoyaltyTesteAtual == "MENOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Red);
                            if (grafico.GraficoFunil3.LoyaltyTesteAtual == "MAIOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Blue);
                            colNameTable++;

                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil3.LoyaltyADV;
                            wworksheet.Column(colNameTable).Width = 20;
                            colNameTable++;
                            currentRow++;

                            colNameTable = 9;

                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil3.BaseAbsAtual;
                            wworksheet.Column(colNameTable).Width = 20;
                            colNameTable++;
                            wworksheet.Cells[currentRow, colNameTable].Value = "";
                            wworksheet.Column(colNameTable).Width = 20;
                            colNameTable++;
                            currentRow++;

                            // Espaço 
                            wworksheet.Cells[currentRow, colNameTable].Value = "";
                            wworksheet.Column(colNameTable).Width = 2;
                            colNameTable++;
                            wworksheet.Cells[currentRow, colNameTable].Value = "";
                            wworksheet.Column(colNameTable).Width = 2;
                            colNameTable++;
                            wworksheet.Cells[currentRow, colNameTable].Value = "";
                            wworksheet.Column(colNameTable).Width = 2;
                            colNameTable++;
                            currentRow++;

                            #endregion

                            #region Coluna 4


                            currentRow = 6;
                            colNameTable = 12;


                            wworksheet.Cells[currentRow - 1, colNameTable].Value = filtro.Onda4.FirstOrDefault().DescItem;

                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil4.ConhecimentoAtual;
                            wworksheet.Column(colNameTable).Width = 20;
                            if (grafico.GraficoFunil4.ConhecimentoTesteAtual == "MENOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Red);
                            if (grafico.GraficoFunil4.ConhecimentoTesteAtual == "MAIOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Blue);
                            colNameTable++;

                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil4.ConhecimentoADV;
                            wworksheet.Column(colNameTable).Width = 20;
                            colNameTable++;
                            currentRow++;

                            colNameTable = 12;

                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil4.ConsideracaoAtual;
                            wworksheet.Column(colNameTable).Width = 20;
                            if (grafico.GraficoFunil4.ConsideracaoTesteAtual == "MENOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Red);
                            if (grafico.GraficoFunil4.ConsideracaoTesteAtual == "MAIOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Blue);
                            colNameTable++;

                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil4.ConsideracaoADV;
                            wworksheet.Column(colNameTable).Width = 20;
                            colNameTable++;
                            currentRow++;

                            colNameTable = 12;

                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil4.UsoAtual;
                            wworksheet.Column(colNameTable).Width = 20;
                            if (grafico.GraficoFunil4.UsoTesteAtual == "MENOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Red);
                            if (grafico.GraficoFunil4.UsoTesteAtual == "MAIOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Blue);
                            colNameTable++;

                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil4.UsoADV;
                            wworksheet.Column(colNameTable).Width = 20;
                            colNameTable++;
                            currentRow++;

                            colNameTable = 12;

                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil4.PreferenciaAtual;
                            wworksheet.Column(colNameTable).Width = 20;
                            if (grafico.GraficoFunil4.PreferenciaTesteAtual == "MENOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Red);
                            if (grafico.GraficoFunil4.PreferenciaTesteAtual == "MAIOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Blue);
                            colNameTable++;

                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil4.PreferenciaADV;
                            wworksheet.Column(colNameTable).Width = 20;
                            colNameTable++;
                            currentRow++;

                            colNameTable = 12;

                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil4.LoyaltyAtual;
                            wworksheet.Column(colNameTable).Width = 20;
                            if (grafico.GraficoFunil4.LoyaltyTesteAtual == "MENOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Red);
                            if (grafico.GraficoFunil4.LoyaltyTesteAtual == "MAIOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Blue);
                            colNameTable++;

                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil4.LoyaltyADV;
                            wworksheet.Column(colNameTable).Width = 20;
                            colNameTable++;
                            currentRow++;

                            colNameTable = 12;

                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil4.BaseAbsAtual;
                            wworksheet.Column(colNameTable).Width = 20;
                            colNameTable++;
                            wworksheet.Cells[currentRow, colNameTable].Value = "";
                            wworksheet.Column(colNameTable).Width = 20;
                            colNameTable++;
                            currentRow++;

                            // Espaço 
                            wworksheet.Cells[currentRow, colNameTable].Value = "";
                            wworksheet.Column(colNameTable).Width = 2;
                            colNameTable++;
                            wworksheet.Cells[currentRow, colNameTable].Value = "";
                            wworksheet.Column(colNameTable).Width = 2;
                            colNameTable++;
                            wworksheet.Cells[currentRow, colNameTable].Value = "";
                            wworksheet.Column(colNameTable).Width = 2;
                            colNameTable++;
                            currentRow++;

                            #endregion



                            currentRow = 7;

                            lastRow++;
                            currentRow = lastRow;
                            colNameTable--;
                            lastColName = colNameTable;

                            //Formata dados
                            //wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 7) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Numberformat.Format = "0.00";
                            //wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 7) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Numberformat.Format = "0.0";
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 7) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Font.Size = 11;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 7) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Font.Name = "Arial Nova Cond";
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 7) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Font.Bold = true;


                        }
                        wworksheet.Select();

                        using (MemoryStream memoryStream = new MemoryStream())
                        {
                            package.SaveAs(memoryStream);

                            memoryStream.Position = 0;

                            returnedMemoryStream = memoryStream;
                        }
                    }
                }

                return File.ReadAllBytes(@"" + fileName);
            }

            catch (Exception)
            {
                throw;
            }
            finally
            {
                File.Delete(fileName);
                File.Delete(templateFileName);
            }
        }

        public byte[] DownloadDashboardThreeComparativoMarcasOLD(GraficoFunilFullLoad grafico, string tituloPlanilha, List<TraducaoComponente> listaTraducaoComponente)
        {
            string fileName = "";
            string templateFileName = "";
            try
            {
                string strDownloadNome;

                strDownloadNome = "Modelo_Exportação_Gráficos_Padrao_Filtro";

                string strPath = AdjustPath(HttpContext.Current.Server.MapPath("~")) + @"Arquivos\ExcelExportacao\workspace\";
                fileName = strPath + strDownloadNome + DateTime.Now.ToString("yyyy - MM - dd HH':'mm':'ss':'fff").Replace(" - ", "").Replace(":", "").Replace(" ", "") + ".xlsx";

                templateFileName = AdjustPath(HttpContext.Current.Server.MapPath("~")) + @"Arquivos\ExcelExportacao\Modelos\" + strDownloadNome + ".xlsx";

                GetNewFileMemory(ref templateFileName);

                if (File.Exists(fileName))
                    File.Delete(fileName);

                MemoryStream returnedMemoryStream = null;

                using (var file = new FileStream(fileName, FileMode.CreateNew))
                {
                    using (var temp = new FileStream(templateFileName, FileMode.Open))
                    {
                        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                        ExcelPackage package = new ExcelPackage(file, temp);

                        ExcelWorksheet wworksheet = package.Workbook.Worksheets["Data"];

                        wworksheet.View.ShowHeaders = true;
                        wworksheet.View.ShowGridLines = false;
                        //Titulo da planilha. Deixar linha 3 vazias formatadas no modelo excel.
                        wworksheet.Cells[3, 2].Value = tituloPlanilha;

                        int currentRow = 6;
                        int colNameTable = 3;
                        int lastColName = 0;


                        if (grafico != null)
                        {
                            lastColName = colNameTable;

                            colNameTable = 2;
                            currentRow++;

                            ////Preenche lateral

                            wworksheet.Cells[currentRow, colNameTable].Value = listaTraducaoComponente.FirstOrDefault(a => a.Objeto.Equals("grafico-funil-conhecimento")).Texto;
                            currentRow++;
                            wworksheet.Cells[currentRow, colNameTable].Value = listaTraducaoComponente.FirstOrDefault(a => a.Objeto.Equals("grafico-funil-consideracao")).Texto;
                            currentRow++;
                            wworksheet.Cells[currentRow, colNameTable].Value = listaTraducaoComponente.FirstOrDefault(a => a.Objeto.Equals("grafico-funil-uso")).Texto;
                            currentRow++;
                            wworksheet.Cells[currentRow, colNameTable].Value = listaTraducaoComponente.FirstOrDefault(a => a.Objeto.Equals("grafico-funil-preferencia")).Texto;
                            currentRow++;
                            wworksheet.Cells[currentRow, colNameTable].Value = "Base";

                            currentRow--;
                            var lastRow = currentRow;
                            colNameTable = 3;
                            currentRow = 7;

                            #region Coluna 1

                            wworksheet.Cells[currentRow - 1, colNameTable].Value = grafico.GraficoFunil1.PeriodoAnt;
                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil1.ConhecimentoAnterior;
                            wworksheet.Column(colNameTable).Width = 20;
                            colNameTable++;
                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil1.ConhecimentoAtual;
                            wworksheet.Column(colNameTable).Width = 20;
                            colNameTable++;
                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil1.ConhecimentoADV;
                            wworksheet.Column(colNameTable).Width = 20;
                            colNameTable++;
                            currentRow++;

                            colNameTable = 3;
                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil1.ConsideracaoAnterior;
                            wworksheet.Column(colNameTable).Width = 20;
                            colNameTable++;
                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil1.ConsideracaoAtual;
                            wworksheet.Column(colNameTable).Width = 20;
                            colNameTable++;
                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil1.ConsideracaoADV;
                            wworksheet.Column(colNameTable).Width = 20;
                            colNameTable++;
                            currentRow++;

                            colNameTable = 3;
                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil1.UsoAnterior;
                            wworksheet.Column(colNameTable).Width = 20;
                            colNameTable++;
                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil1.UsoAtual;
                            wworksheet.Column(colNameTable).Width = 20;
                            colNameTable++;
                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil1.UsoADV;
                            wworksheet.Column(colNameTable).Width = 20;
                            colNameTable++;
                            currentRow++;

                            colNameTable = 3;
                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil1.PreferenciaAnterior;
                            wworksheet.Column(colNameTable).Width = 20;
                            colNameTable++;
                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil1.PreferenciaAtual;
                            wworksheet.Column(colNameTable).Width = 20;
                            colNameTable++;
                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil1.PreferenciaADV;
                            wworksheet.Column(colNameTable).Width = 20;
                            colNameTable++;
                            currentRow++;

                            colNameTable = 3;
                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil1.BaseAbsAnterior;
                            wworksheet.Column(colNameTable).Width = 20;
                            colNameTable++;
                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil1.BaseAbsAtual;
                            wworksheet.Column(colNameTable).Width = 20;
                            colNameTable++;
                            wworksheet.Cells[currentRow, colNameTable].Value = "";
                            wworksheet.Column(colNameTable).Width = 20;
                            colNameTable++;
                            currentRow++;

                            // Espaço 
                            wworksheet.Cells[currentRow, colNameTable].Value = "";
                            wworksheet.Column(colNameTable).Width = 2;
                            colNameTable++;
                            wworksheet.Cells[currentRow, colNameTable].Value = "";
                            wworksheet.Column(colNameTable).Width = 2;
                            colNameTable++;
                            wworksheet.Cells[currentRow, colNameTable].Value = "";
                            wworksheet.Column(colNameTable).Width = 2;
                            colNameTable++;
                            currentRow++;

                            #endregion

                            #region Coluna 2


                            currentRow = 7;
                            colNameTable = 7;


                            wworksheet.Cells[currentRow - 1, colNameTable].Value = grafico.GraficoFunil2.PeriodoAnt;
                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil2.ConhecimentoAnterior;
                            wworksheet.Column(colNameTable).Width = 20;
                            colNameTable++;
                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil2.ConhecimentoAtual;
                            wworksheet.Column(colNameTable).Width = 20;
                            colNameTable++;
                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil2.ConhecimentoADV;
                            wworksheet.Column(colNameTable).Width = 20;
                            colNameTable++;
                            currentRow++;

                            colNameTable = 7;
                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil2.ConsideracaoAnterior;
                            wworksheet.Column(colNameTable).Width = 20;
                            colNameTable++;
                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil2.ConsideracaoAtual;
                            wworksheet.Column(colNameTable).Width = 20;
                            colNameTable++;
                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil2.ConsideracaoADV;
                            wworksheet.Column(colNameTable).Width = 20;
                            colNameTable++;
                            currentRow++;

                            colNameTable = 7;
                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil2.UsoAnterior;
                            wworksheet.Column(colNameTable).Width = 20;
                            colNameTable++;
                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil2.UsoAtual;
                            wworksheet.Column(colNameTable).Width = 20;
                            colNameTable++;
                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil2.UsoADV;
                            wworksheet.Column(colNameTable).Width = 20;
                            colNameTable++;
                            currentRow++;

                            colNameTable = 7;
                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil2.PreferenciaAnterior;
                            wworksheet.Column(colNameTable).Width = 20;
                            colNameTable++;
                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil2.PreferenciaAtual;
                            wworksheet.Column(colNameTable).Width = 20;
                            colNameTable++;
                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil2.PreferenciaADV;
                            wworksheet.Column(colNameTable).Width = 20;
                            colNameTable++;
                            currentRow++;

                            colNameTable = 7;
                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil2.BaseAbsAnterior;
                            wworksheet.Column(colNameTable).Width = 20;
                            colNameTable++;
                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil2.BaseAbsAtual;
                            wworksheet.Column(colNameTable).Width = 20;
                            colNameTable++;
                            wworksheet.Cells[currentRow, colNameTable].Value = "";
                            wworksheet.Column(colNameTable).Width = 20;
                            colNameTable++;
                            currentRow++;

                            // Espaço 
                            wworksheet.Cells[currentRow, colNameTable].Value = "";
                            wworksheet.Column(colNameTable).Width = 2;
                            colNameTable++;
                            wworksheet.Cells[currentRow, colNameTable].Value = "";
                            wworksheet.Column(colNameTable).Width = 2;
                            colNameTable++;
                            wworksheet.Cells[currentRow, colNameTable].Value = "";
                            wworksheet.Column(colNameTable).Width = 2;
                            colNameTable++;
                            currentRow++;

                            #endregion

                            #region Coluna 3


                            currentRow = 7;
                            colNameTable = 11;


                            wworksheet.Cells[currentRow - 1, colNameTable].Value = grafico.GraficoFunil3.PeriodoAnt;
                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil3.ConhecimentoAnterior;
                            wworksheet.Column(colNameTable).Width = 20;
                            colNameTable++;
                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil3.ConhecimentoAtual;
                            wworksheet.Column(colNameTable).Width = 20;
                            colNameTable++;
                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil3.ConhecimentoADV;
                            wworksheet.Column(colNameTable).Width = 20;
                            colNameTable++;
                            currentRow++;

                            colNameTable = 11;
                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil3.ConsideracaoAnterior;
                            wworksheet.Column(colNameTable).Width = 20;
                            colNameTable++;
                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil3.ConsideracaoAtual;
                            wworksheet.Column(colNameTable).Width = 20;
                            colNameTable++;
                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil3.ConsideracaoADV;
                            wworksheet.Column(colNameTable).Width = 20;
                            colNameTable++;
                            currentRow++;

                            colNameTable = 11;
                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil3.UsoAnterior;
                            wworksheet.Column(colNameTable).Width = 20;
                            colNameTable++;
                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil3.UsoAtual;
                            wworksheet.Column(colNameTable).Width = 20;
                            colNameTable++;
                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil3.UsoADV;
                            wworksheet.Column(colNameTable).Width = 20;
                            colNameTable++;
                            currentRow++;

                            colNameTable = 11;
                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil3.PreferenciaAnterior;
                            wworksheet.Column(colNameTable).Width = 20;
                            colNameTable++;
                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil3.PreferenciaAtual;
                            wworksheet.Column(colNameTable).Width = 20;
                            colNameTable++;
                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil3.PreferenciaADV;
                            wworksheet.Column(colNameTable).Width = 20;
                            colNameTable++;
                            currentRow++;

                            colNameTable = 11;
                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil3.BaseAbsAnterior;
                            wworksheet.Column(colNameTable).Width = 20;
                            colNameTable++;
                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil3.BaseAbsAtual;
                            wworksheet.Column(colNameTable).Width = 20;
                            colNameTable++;
                            wworksheet.Cells[currentRow, colNameTable].Value = "";
                            wworksheet.Column(colNameTable).Width = 20;
                            colNameTable++;
                            currentRow++;

                            // Espaço 
                            wworksheet.Cells[currentRow, colNameTable].Value = "";
                            wworksheet.Column(colNameTable).Width = 2;
                            colNameTable++;
                            wworksheet.Cells[currentRow, colNameTable].Value = "";
                            wworksheet.Column(colNameTable).Width = 2;
                            colNameTable++;
                            wworksheet.Cells[currentRow, colNameTable].Value = "";
                            wworksheet.Column(colNameTable).Width = 2;
                            colNameTable++;
                            currentRow++;

                            #endregion

                            #region Coluna 4


                            currentRow = 7;
                            colNameTable = 15;

                            wworksheet.Cells[currentRow - 1, colNameTable].Value = grafico.GraficoFunil4.PeriodoAnt;
                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil4.ConhecimentoAnterior;
                            wworksheet.Column(colNameTable).Width = 20;
                            colNameTable++;
                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil4.ConhecimentoAtual;
                            wworksheet.Column(colNameTable).Width = 20;
                            colNameTable++;
                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil4.ConhecimentoADV;
                            wworksheet.Column(colNameTable).Width = 20;
                            colNameTable++;
                            currentRow++;

                            colNameTable = 15;
                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil4.ConsideracaoAnterior;
                            wworksheet.Column(colNameTable).Width = 20;
                            colNameTable++;
                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil4.ConsideracaoAtual;
                            wworksheet.Column(colNameTable).Width = 20;
                            colNameTable++;
                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil4.ConsideracaoADV;
                            wworksheet.Column(colNameTable).Width = 20;
                            colNameTable++;
                            currentRow++;

                            colNameTable = 15;
                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil4.UsoAnterior;
                            wworksheet.Column(colNameTable).Width = 20;
                            colNameTable++;
                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil4.UsoAtual;
                            wworksheet.Column(colNameTable).Width = 20;
                            colNameTable++;
                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil4.UsoADV;
                            wworksheet.Column(colNameTable).Width = 20;
                            colNameTable++;
                            currentRow++;

                            colNameTable = 15;
                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil4.PreferenciaAnterior;
                            wworksheet.Column(colNameTable).Width = 20;
                            colNameTable++;
                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil4.PreferenciaAtual;
                            wworksheet.Column(colNameTable).Width = 20;
                            colNameTable++;
                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil4.PreferenciaADV;
                            wworksheet.Column(colNameTable).Width = 20;
                            colNameTable++;
                            currentRow++;

                            colNameTable = 15;
                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil4.BaseAbsAnterior;
                            wworksheet.Column(colNameTable).Width = 20;
                            colNameTable++;
                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil4.BaseAbsAtual;
                            wworksheet.Column(colNameTable).Width = 20;
                            colNameTable++;
                            wworksheet.Cells[currentRow, colNameTable].Value = "";
                            wworksheet.Column(colNameTable).Width = 20;
                            colNameTable++;
                            currentRow++;

                            // Espaço 
                            wworksheet.Cells[currentRow, colNameTable].Value = "";
                            wworksheet.Column(colNameTable).Width = 2;
                            colNameTable++;
                            wworksheet.Cells[currentRow, colNameTable].Value = "";
                            wworksheet.Column(colNameTable).Width = 2;
                            colNameTable++;
                            wworksheet.Cells[currentRow, colNameTable].Value = "";
                            wworksheet.Column(colNameTable).Width = 2;
                            colNameTable++;
                            currentRow++;

                            #endregion


                            colNameTable = 2;
                            currentRow = 15;

                            ////Preenche lateral

                            wworksheet.Cells[currentRow, colNameTable].Value = listaTraducaoComponente.FirstOrDefault(a => a.Objeto.Equals("grafico-funil-conhecimento")).Texto;
                            currentRow++;
                            wworksheet.Cells[currentRow, colNameTable].Value = listaTraducaoComponente.FirstOrDefault(a => a.Objeto.Equals("grafico-funil-consideracao")).Texto;
                            currentRow++;
                            wworksheet.Cells[currentRow, colNameTable].Value = listaTraducaoComponente.FirstOrDefault(a => a.Objeto.Equals("grafico-funil-uso")).Texto;
                            currentRow++;
                            wworksheet.Cells[currentRow, colNameTable].Value = listaTraducaoComponente.FirstOrDefault(a => a.Objeto.Equals("grafico-funil-preferencia")).Texto;
                            currentRow++;
                            wworksheet.Cells[currentRow, colNameTable].Value = "Base";


                            #region Coluna 5


                            currentRow = 15;
                            colNameTable = 3;

                            wworksheet.Cells[currentRow - 1, colNameTable].Value = grafico.GraficoFunil5.PeriodoAnt;
                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil5.ConhecimentoAnterior;
                            wworksheet.Column(colNameTable).Width = 20;
                            colNameTable++;
                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil5.ConhecimentoAtual;
                            wworksheet.Column(colNameTable).Width = 20;
                            colNameTable++;
                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil5.ConhecimentoADV;
                            wworksheet.Column(colNameTable).Width = 20;
                            colNameTable++;
                            currentRow++;

                            colNameTable = 3;
                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil5.ConsideracaoAnterior;
                            wworksheet.Column(colNameTable).Width = 20;
                            colNameTable++;
                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil5.ConsideracaoAtual;
                            wworksheet.Column(colNameTable).Width = 20;
                            colNameTable++;
                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil5.ConsideracaoADV;
                            wworksheet.Column(colNameTable).Width = 20;
                            colNameTable++;
                            currentRow++;

                            colNameTable = 3;
                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil5.UsoAnterior;
                            wworksheet.Column(colNameTable).Width = 20;
                            colNameTable++;
                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil5.UsoAtual;
                            wworksheet.Column(colNameTable).Width = 20;
                            colNameTable++;
                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil5.UsoADV;
                            wworksheet.Column(colNameTable).Width = 20;
                            colNameTable++;
                            currentRow++;

                            colNameTable = 3;
                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil5.PreferenciaAnterior;
                            wworksheet.Column(colNameTable).Width = 20;
                            colNameTable++;
                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil5.PreferenciaAtual;
                            wworksheet.Column(colNameTable).Width = 20;
                            colNameTable++;
                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil5.PreferenciaADV;
                            wworksheet.Column(colNameTable).Width = 20;
                            colNameTable++;
                            currentRow++;

                            colNameTable = 3;
                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil5.BaseAbsAnterior;
                            wworksheet.Column(colNameTable).Width = 20;
                            colNameTable++;
                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil5.BaseAbsAtual;
                            wworksheet.Column(colNameTable).Width = 20;
                            colNameTable++;
                            wworksheet.Cells[currentRow, colNameTable].Value = "";
                            wworksheet.Column(colNameTable).Width = 20;
                            colNameTable++;
                            currentRow++;

                            // Espaço 
                            wworksheet.Cells[currentRow, colNameTable].Value = "";
                            wworksheet.Column(colNameTable).Width = 2;
                            colNameTable++;
                            wworksheet.Cells[currentRow, colNameTable].Value = "";
                            wworksheet.Column(colNameTable).Width = 2;
                            colNameTable++;
                            wworksheet.Cells[currentRow, colNameTable].Value = "";
                            wworksheet.Column(colNameTable).Width = 2;
                            colNameTable++;
                            currentRow++;

                            #endregion

                            #region Coluna 6


                            currentRow = 15;
                            colNameTable = 7;

                            wworksheet.Cells[currentRow - 1, colNameTable].Value = grafico.GraficoFunil6.PeriodoAnt;
                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil6.ConhecimentoAnterior;
                            wworksheet.Column(colNameTable).Width = 20;
                            colNameTable++;
                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil6.ConhecimentoAtual;
                            wworksheet.Column(colNameTable).Width = 20;
                            colNameTable++;
                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil6.ConhecimentoADV;
                            wworksheet.Column(colNameTable).Width = 20;
                            colNameTable++;
                            currentRow++;

                            colNameTable = 7;
                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil6.ConsideracaoAnterior;
                            wworksheet.Column(colNameTable).Width = 20;
                            colNameTable++;
                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil6.ConsideracaoAtual;
                            wworksheet.Column(colNameTable).Width = 20;
                            colNameTable++;
                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil6.ConsideracaoADV;
                            wworksheet.Column(colNameTable).Width = 20;
                            colNameTable++;
                            currentRow++;

                            colNameTable = 7;
                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil6.UsoAnterior;
                            wworksheet.Column(colNameTable).Width = 20;
                            colNameTable++;
                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil6.UsoAtual;
                            wworksheet.Column(colNameTable).Width = 20;
                            colNameTable++;
                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil6.UsoADV;
                            wworksheet.Column(colNameTable).Width = 20;
                            colNameTable++;
                            currentRow++;

                            colNameTable = 7;
                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil6.PreferenciaAnterior;
                            wworksheet.Column(colNameTable).Width = 20;
                            colNameTable++;
                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil6.PreferenciaAtual;
                            wworksheet.Column(colNameTable).Width = 20;
                            colNameTable++;
                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil6.PreferenciaADV;
                            wworksheet.Column(colNameTable).Width = 20;
                            colNameTable++;
                            currentRow++;

                            colNameTable = 7;
                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil6.BaseAbsAnterior;
                            wworksheet.Column(colNameTable).Width = 20;
                            colNameTable++;
                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil6.BaseAbsAtual;
                            wworksheet.Column(colNameTable).Width = 20;
                            colNameTable++;
                            wworksheet.Cells[currentRow, colNameTable].Value = "";
                            wworksheet.Column(colNameTable).Width = 20;
                            colNameTable++;
                            currentRow++;

                            // Espaço 
                            wworksheet.Cells[currentRow, colNameTable].Value = "";
                            wworksheet.Column(colNameTable).Width = 2;
                            colNameTable++;
                            wworksheet.Cells[currentRow, colNameTable].Value = "";
                            wworksheet.Column(colNameTable).Width = 2;
                            colNameTable++;
                            wworksheet.Cells[currentRow, colNameTable].Value = "";
                            wworksheet.Column(colNameTable).Width = 2;
                            colNameTable++;
                            currentRow++;

                            #endregion

                            #region Coluna 7


                            currentRow = 15;
                            colNameTable = 11;

                            wworksheet.Cells[currentRow - 1, colNameTable].Value = grafico.GraficoFunil7.PeriodoAnt;
                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil7.ConhecimentoAnterior;
                            wworksheet.Column(colNameTable).Width = 20;
                            colNameTable++;
                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil7.ConhecimentoAtual;
                            wworksheet.Column(colNameTable).Width = 20;
                            colNameTable++;
                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil7.ConhecimentoADV;
                            wworksheet.Column(colNameTable).Width = 20;
                            colNameTable++;
                            currentRow++;

                            colNameTable = 11;
                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil7.ConsideracaoAnterior;
                            wworksheet.Column(colNameTable).Width = 20;
                            colNameTable++;
                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil7.ConsideracaoAtual;
                            wworksheet.Column(colNameTable).Width = 20;
                            colNameTable++;
                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil7.ConsideracaoADV;
                            wworksheet.Column(colNameTable).Width = 20;
                            colNameTable++;
                            currentRow++;

                            colNameTable = 11;
                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil7.UsoAnterior;
                            wworksheet.Column(colNameTable).Width = 20;
                            colNameTable++;
                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil7.UsoAtual;
                            wworksheet.Column(colNameTable).Width = 20;
                            colNameTable++;
                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil7.UsoADV;
                            wworksheet.Column(colNameTable).Width = 20;
                            colNameTable++;
                            currentRow++;

                            colNameTable = 11;
                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil7.PreferenciaAnterior;
                            wworksheet.Column(colNameTable).Width = 20;
                            colNameTable++;
                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil7.PreferenciaAtual;
                            wworksheet.Column(colNameTable).Width = 20;
                            colNameTable++;
                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil7.PreferenciaADV;
                            wworksheet.Column(colNameTable).Width = 20;
                            colNameTable++;
                            currentRow++;

                            colNameTable = 11;
                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil7.BaseAbsAnterior;
                            wworksheet.Column(colNameTable).Width = 20;
                            colNameTable++;
                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil7.BaseAbsAtual;
                            wworksheet.Column(colNameTable).Width = 20;
                            colNameTable++;
                            wworksheet.Cells[currentRow, colNameTable].Value = "";
                            wworksheet.Column(colNameTable).Width = 20;
                            colNameTable++;
                            currentRow++;

                            // Espaço 
                            wworksheet.Cells[currentRow, colNameTable].Value = "";
                            wworksheet.Column(colNameTable).Width = 2;
                            colNameTable++;
                            wworksheet.Cells[currentRow, colNameTable].Value = "";
                            wworksheet.Column(colNameTable).Width = 2;
                            colNameTable++;
                            wworksheet.Cells[currentRow, colNameTable].Value = "";
                            wworksheet.Column(colNameTable).Width = 2;
                            colNameTable++;
                            currentRow++;

                            #endregion

                            #region Coluna 8


                            currentRow = 15;
                            colNameTable = 15;

                            wworksheet.Cells[currentRow - 1, colNameTable].Value = grafico.GraficoFunil8.PeriodoAnt;
                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil8.ConhecimentoAnterior;
                            wworksheet.Column(colNameTable).Width = 20;
                            colNameTable++;
                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil8.ConhecimentoAtual;
                            wworksheet.Column(colNameTable).Width = 20;
                            colNameTable++;
                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil8.ConhecimentoADV;
                            wworksheet.Column(colNameTable).Width = 20;
                            colNameTable++;
                            currentRow++;

                            colNameTable = 15;
                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil8.ConsideracaoAnterior;
                            wworksheet.Column(colNameTable).Width = 20;
                            colNameTable++;
                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil8.ConsideracaoAtual;
                            wworksheet.Column(colNameTable).Width = 20;
                            colNameTable++;
                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil8.ConsideracaoADV;
                            wworksheet.Column(colNameTable).Width = 20;
                            colNameTable++;
                            currentRow++;

                            colNameTable = 15;
                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil8.UsoAnterior;
                            wworksheet.Column(colNameTable).Width = 20;
                            colNameTable++;
                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil8.UsoAtual;
                            wworksheet.Column(colNameTable).Width = 20;
                            colNameTable++;
                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil8.UsoADV;
                            wworksheet.Column(colNameTable).Width = 20;
                            colNameTable++;
                            currentRow++;

                            colNameTable = 15;
                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil8.PreferenciaAnterior;
                            wworksheet.Column(colNameTable).Width = 20;
                            colNameTable++;
                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil8.PreferenciaAtual;
                            wworksheet.Column(colNameTable).Width = 20;
                            colNameTable++;
                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil8.PreferenciaADV;
                            wworksheet.Column(colNameTable).Width = 20;
                            colNameTable++;
                            currentRow++;

                            colNameTable = 15;
                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil8.BaseAbsAnterior;
                            wworksheet.Column(colNameTable).Width = 20;
                            colNameTable++;
                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoFunil8.BaseAbsAtual;
                            wworksheet.Column(colNameTable).Width = 20;
                            colNameTable++;
                            wworksheet.Cells[currentRow, colNameTable].Value = "";
                            wworksheet.Column(colNameTable).Width = 20;
                            colNameTable++;
                            currentRow++;

                            // Espaço 
                            wworksheet.Cells[currentRow, colNameTable].Value = "";
                            wworksheet.Column(colNameTable).Width = 2;
                            colNameTable++;
                            wworksheet.Cells[currentRow, colNameTable].Value = "";
                            wworksheet.Column(colNameTable).Width = 2;
                            colNameTable++;
                            wworksheet.Cells[currentRow, colNameTable].Value = "";
                            wworksheet.Column(colNameTable).Width = 2;
                            colNameTable++;
                            currentRow++;

                            #endregion

                            currentRow = 7;

                            lastRow++;
                            currentRow = lastRow;
                            colNameTable--;
                            lastColName = colNameTable;

                            //Formata dados
                            //wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 7) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Numberformat.Format = "0.00";
                            //wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 7) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Numberformat.Format = "0.0";
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 7) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Font.Size = 11;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 7) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Font.Name = "Arial Nova Cond";
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 7) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Font.Bold = true;


                        }
                        wworksheet.Select();

                        using (MemoryStream memoryStream = new MemoryStream())
                        {
                            package.SaveAs(memoryStream);

                            memoryStream.Position = 0;

                            returnedMemoryStream = memoryStream;
                        }
                    }
                }

                return File.ReadAllBytes(@"" + fileName);
            }

            catch (Exception)
            {
                throw;
            }
            finally
            {
                File.Delete(fileName);
                File.Delete(templateFileName);
            }
        }

        #endregion


        #region DownloadDashboardFour
        public byte[] DownloadDashboardFourComparativoMarcas(List<GraficoAranhaRetornoProcedure> grafico, string tituloPlanilha, List<TraducaoComponente> listaTraducaoComponente)
        {
            string fileName = "";
            string templateFileName = "";
            try
            {
                string strDownloadNome;

                strDownloadNome = "Modelo_Exportação_Gráficos_Padrao_Four";

                string strPath = AdjustPath(HttpContext.Current.Server.MapPath("~")) + @"Arquivos\ExcelExportacao\workspace\";
                fileName = strPath + strDownloadNome + DateTime.Now.ToString("yyyy - MM - dd HH':'mm':'ss':'fff").Replace(" - ", "").Replace(":", "").Replace(" ", "") + ".xlsx";

                templateFileName = AdjustPath(HttpContext.Current.Server.MapPath("~")) + @"Arquivos\ExcelExportacao\Modelos\" + strDownloadNome + ".xlsx";

                GetNewFileMemory(ref templateFileName);

                if (File.Exists(fileName))
                    File.Delete(fileName);

                MemoryStream returnedMemoryStream = null;

                using (var file = new FileStream(fileName, FileMode.CreateNew))
                {
                    using (var temp = new FileStream(templateFileName, FileMode.Open))
                    {
                        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                        ExcelPackage package = new ExcelPackage(file, temp);

                        ExcelWorksheet wworksheet = package.Workbook.Worksheets["Data"];

                        wworksheet.View.ShowHeaders = true;
                        wworksheet.View.ShowGridLines = false;
                        //Titulo da planilha. Deixar linha 3 vazias formatadas no modelo excel.
                        wworksheet.Cells[3, 2].Value = tituloPlanilha;

                        int currentRow = 6;
                        int colNameTable = 3;
                        int lastColName = 0;


                        if (grafico != null)
                        {

                            // Preenche cabeçalho com os periodos:
                            foreach (var item in grafico.Select(a => a.DescMarca).Distinct().ToList())
                            {
                                wworksheet.Cells[currentRow, colNameTable].Value = item;
                                colNameTable++;
                            }



                            colNameTable--;
                            lastColName = colNameTable;

                            //Formata cabeçalho
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(3, 6) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Font.Size = 11;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(3, 6) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Font.Name = "Arial Nova Cond";
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(3, 6) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Font.Bold = true;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(3, 6) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(3, 6) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(3, 6) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(3, 6) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(3, 6) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(3, 6) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                            colNameTable = 2;
                            currentRow++;

                            ////Preenche lateral
                            foreach (var item in grafico.Select(a => a.Descricao).Distinct().ToList())
                            {
                                wworksheet.Cells[currentRow, colNameTable].Value = item;
                                currentRow++;
                            }
                            wworksheet.Cells[currentRow, colNameTable].Value = "Base";
                            currentRow++;

                            currentRow--;
                            var lastRow = currentRow;
                            colNameTable = 3;
                            currentRow = 7;

                            foreach (var item in grafico.Select(a => a.Descricao).Distinct().ToList())
                            {
                                colNameTable = 3;
                                var itemLinha = grafico.Where(a => a.Descricao == item).ToList();

                                // MONTA COLUNAS DADOS
                                foreach (var dadoLinha in itemLinha)
                                {
                                    wworksheet.Cells[currentRow, colNameTable].Value = dadoLinha.Perc;
                                    colNameTable++;
                                }


                                currentRow++;
                            }

                            var ultimaLinha = currentRow;


                            foreach (var item in grafico.Select(a => a.Descricao).Distinct().Take(1).ToList())
                            {
                                colNameTable = 3;
                                var itemLinha = grafico.Where(a => a.Descricao == item).ToList();

                                // MONTA COLUNAS DADOS
                                foreach (var dadoLinha in itemLinha)
                                {
                                    wworksheet.Cells[currentRow, colNameTable].Value = dadoLinha.Base;
                                    colNameTable++;
                                }


                                currentRow++;
                            }



                            currentRow = 7;

                            lastRow++;
                            currentRow = lastRow;
                            colNameTable--;
                            lastColName = colNameTable;

                            //Formata dados
                            //wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 7) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Numberformat.Format = "0.00";
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 7) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Font.Size = 11;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 7) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Font.Name = "Arial Nova Cond";
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 7) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Font.Bold = true;

                            // Preenche lateral
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 11) + ":" + ExcelUtil.GetCelulaByColumnIndex(3, 11)].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 11) + ":" + ExcelUtil.GetCelulaByColumnIndex(3, 11)].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 11) + ":" + ExcelUtil.GetCelulaByColumnIndex(3, 11)].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 11) + ":" + ExcelUtil.GetCelulaByColumnIndex(3, 11)].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 11) + ":" + ExcelUtil.GetCelulaByColumnIndex(3, 11)].Style.Border.Top.Style = ExcelBorderStyle.Thin;

                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 7) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 7) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 7) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 7) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 7) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                        }
                        wworksheet.Select();

                        using (MemoryStream memoryStream = new MemoryStream())
                        {
                            package.SaveAs(memoryStream);

                            memoryStream.Position = 0;

                            returnedMemoryStream = memoryStream;
                        }
                    }
                }

                return File.ReadAllBytes(@"" + fileName);
            }

            catch (Exception)
            {
                throw;
            }
            finally
            {
                File.Delete(fileName);
                File.Delete(templateFileName);
            }
        }


        public byte[] DownloadDashboardFourGraficoLinhasImagem(GraficoLinhasImagem grafico, string tituloPlanilha, List<TraducaoComponente> listaTraducaoComponente, FiltroPadraoExcel filtro)
        {
            string fileName = "";
            string templateFileName = "";
            try
            {
                string strDownloadNome;

                strDownloadNome = "Modelo_Exportação_Gráficos_Padrao_linhas";

                string strPath = AdjustPath(HttpContext.Current.Server.MapPath("~")) + @"Arquivos\ExcelExportacao\workspace\";
                fileName = strPath + strDownloadNome + DateTime.Now.ToString("yyyy - MM - dd HH':'mm':'ss':'fff").Replace(" - ", "").Replace(":", "").Replace(" ", "") + ".xlsx";

                templateFileName = AdjustPath(HttpContext.Current.Server.MapPath("~")) + @"Arquivos\ExcelExportacao\Modelos\" + strDownloadNome + ".xlsx";

                GetNewFileMemory(ref templateFileName);

                if (File.Exists(fileName))
                    File.Delete(fileName);

                MemoryStream returnedMemoryStream = null;

                using (var file = new FileStream(fileName, FileMode.CreateNew))
                {
                    using (var temp = new FileStream(templateFileName, FileMode.Open))
                    {
                        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                        ExcelPackage package = new ExcelPackage(file, temp);

                        ExcelWorksheet wworksheet = package.Workbook.Worksheets["Data"];

                        wworksheet.View.ShowHeaders = true;
                        wworksheet.View.ShowGridLines = false;
                        //Titulo da planilha. Deixar linha 3 vazias formatadas no modelo excel.
                        wworksheet.Cells[3, 2].Value = tituloPlanilha;

                        int currentRow = 6;
                        int colNameTable = 3;
                        int lastColName = 0;


                        if (grafico != null)
                        {

                            // Preenche cabeçalho com os periodos:



                            wworksheet.Cells[currentRow - 1, colNameTable].Value = filtro.Marca1.FirstOrDefault().DescItem;
                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoLinhasImagem1.FirstOrDefault().DescMarca;
                            colNameTable++;
                            wworksheet.Cells[currentRow, colNameTable].Value = "";
                            colNameTable++;

                            wworksheet.Cells[currentRow - 1, colNameTable].Value = filtro.Marca2.FirstOrDefault().DescItem;
                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoLinhasImagem2.FirstOrDefault().DescMarca;
                            colNameTable++;
                            wworksheet.Cells[currentRow, colNameTable].Value = "";
                            colNameTable++;

                            wworksheet.Cells[currentRow - 1, colNameTable].Value = filtro.Marca3.FirstOrDefault().DescItem;
                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoLinhasImagem3.FirstOrDefault().DescMarca;
                            colNameTable++;
                            wworksheet.Cells[currentRow, colNameTable].Value = "";
                            colNameTable++;

                            wworksheet.Cells[currentRow - 1, colNameTable].Value = filtro.Marca4.FirstOrDefault().DescItem;
                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoLinhasImagem4.FirstOrDefault().DescMarca;
                            colNameTable++;
                            wworksheet.Cells[currentRow, colNameTable].Value = "";
                            colNameTable++;

                            wworksheet.Cells[currentRow - 1, colNameTable].Value = filtro.Marca5.FirstOrDefault().DescItem;
                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoLinhasImagem5.FirstOrDefault().DescMarca;
                            colNameTable++;
                            wworksheet.Cells[currentRow, colNameTable].Value = "";
                            colNameTable++;


                            colNameTable = 3;


                            for (int i = 0; i < 5; i++)
                            {
                                wworksheet.Cells[currentRow, colNameTable].Value = "Anterior";
                                colNameTable++;
                                wworksheet.Cells[currentRow, colNameTable].Value = "Atual";
                                colNameTable++;

                            }


                            colNameTable--;
                            lastColName = colNameTable;

                            //Formata cabeçalho
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(3, 6) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Font.Size = 11;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(3, 6) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Font.Name = "Arial Nova Cond";
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(3, 6) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Font.Bold = true;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(3, 6) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(3, 6) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(3, 6) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(3, 6) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(3, 6) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(3, 6) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                            colNameTable = 2;
                            currentRow++;

                            ////Preenche lateral
                            foreach (var item in grafico.GraficoLinhasImagem1.Select(a => a.Descricao).Distinct().ToList())
                            {
                                wworksheet.Cells[currentRow, colNameTable].Value = item;
                                currentRow++;
                            }

                            wworksheet.Cells[currentRow, colNameTable].Value = "Base";
                            currentRow++;



                            currentRow--;
                            var lastRow = currentRow;
                            colNameTable = 3;
                            currentRow = 7;

                            foreach (var linha in grafico.GraficoLinhasImagem1)
                            {
                                colNameTable = 3;
                                wworksheet.Cells[currentRow, colNameTable].Value = linha.PercAnterior;
                                colNameTable++;
                                wworksheet.Cells[currentRow, colNameTable].Value = linha.PercAtual;
                                if (linha.TestSig == "MENOR")
                                    wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Red);
                                if (linha.TestSig == "MAIOR")
                                    wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Blue);
                                colNameTable++;
                                currentRow++;
                            }

                            colNameTable = 3;
                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoLinhasImagem1[0].BaseAnt;
                            colNameTable++;
                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoLinhasImagem1[0].Base;

                            colNameTable++;
                            currentRow++;


                            currentRow = 7;

                            foreach (var linha in grafico.GraficoLinhasImagem2)
                            {
                                colNameTable = 5;
                                wworksheet.Cells[currentRow, colNameTable].Value = linha.PercAnterior;
                                colNameTable++;
                                wworksheet.Cells[currentRow, colNameTable].Value = linha.PercAtual;
                                if (linha.TestSig == "MENOR")
                                    wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Red);
                                if (linha.TestSig == "MAIOR")
                                    wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Blue);
                                colNameTable++;
                                currentRow++;
                            }

                            colNameTable = 5;
                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoLinhasImagem2[0].BaseAnt;
                            colNameTable++;
                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoLinhasImagem2[0].Base;

                            colNameTable++;
                            currentRow++;

                            currentRow = 7;
                            foreach (var linha in grafico.GraficoLinhasImagem3)
                            {
                                colNameTable = 7;
                                wworksheet.Cells[currentRow, colNameTable].Value = linha.PercAnterior;
                                colNameTable++;
                                wworksheet.Cells[currentRow, colNameTable].Value = linha.PercAtual;
                                if (linha.TestSig == "MENOR")
                                    wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Red);
                                if (linha.TestSig == "MAIOR")
                                    wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Blue);
                                colNameTable++;
                                currentRow++;
                            }
                            colNameTable = 7;
                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoLinhasImagem3[0].BaseAnt;
                            colNameTable++;
                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoLinhasImagem3[0].Base;

                            colNameTable++;
                            currentRow++;


                            currentRow = 7;
                            foreach (var linha in grafico.GraficoLinhasImagem4)
                            {
                                colNameTable = 9;
                                wworksheet.Cells[currentRow, colNameTable].Value = linha.PercAnterior;
                                colNameTable++;
                                wworksheet.Cells[currentRow, colNameTable].Value = linha.PercAtual;
                                if (linha.TestSig == "MENOR")
                                    wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Red);
                                if (linha.TestSig == "MAIOR")
                                    wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Blue);
                                colNameTable++;
                                currentRow++;
                            }
                            colNameTable = 9;
                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoLinhasImagem4[0].BaseAnt;
                            colNameTable++;
                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoLinhasImagem4[0].Base;

                            colNameTable++;
                            currentRow++;


                            currentRow = 7;
                            foreach (var linha in grafico.GraficoLinhasImagem5)
                            {
                                colNameTable = 11;
                                wworksheet.Cells[currentRow, colNameTable].Value = linha.PercAnterior;
                                colNameTable++;
                                wworksheet.Cells[currentRow, colNameTable].Value = linha.PercAtual;
                                if (linha.TestSig == "MENOR")
                                    wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Red);
                                if (linha.TestSig == "MAIOR")
                                    wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Blue);
                                colNameTable++;
                                currentRow++;
                            }

                            colNameTable = 11;
                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoLinhasImagem5[0].BaseAnt;
                            colNameTable++;
                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoLinhasImagem5[0].Base;

                            colNameTable++;
                            currentRow++;




                            //foreach (var item in grafico.Select(a => a.Descricao).Distinct().ToList())
                            //{
                            //    colNameTable = 3;
                            //    var itemLinha = grafico.Where(a => a.Descricao == item).ToList();

                            //    // MONTA COLUNAS DADOS
                            //    foreach (var dadoLinha in itemLinha)
                            //    {
                            //        wworksheet.Cells[currentRow, colNameTable].Value = dadoLinha.Perc;
                            //        colNameTable++;
                            //    }

                            //    currentRow++;
                            //}

                            currentRow = 7;

                            lastRow++;
                            currentRow = lastRow;
                            colNameTable--;
                            lastColName = colNameTable;

                            //Formata dados
                            //wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 7) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Numberformat.Format = "0.0";
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 7) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Font.Size = 11;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 7) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Font.Name = "Arial Nova Cond";
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 7) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Font.Bold = true;

                            // Preenche lateral
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 11) + ":" + ExcelUtil.GetCelulaByColumnIndex(3, 11)].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 11) + ":" + ExcelUtil.GetCelulaByColumnIndex(3, 11)].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 11) + ":" + ExcelUtil.GetCelulaByColumnIndex(3, 11)].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 11) + ":" + ExcelUtil.GetCelulaByColumnIndex(3, 11)].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 11) + ":" + ExcelUtil.GetCelulaByColumnIndex(3, 11)].Style.Border.Top.Style = ExcelBorderStyle.Thin;

                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 7) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 7) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 7) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 7) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 7) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                        }
                        wworksheet.Select();

                        using (MemoryStream memoryStream = new MemoryStream())
                        {
                            package.SaveAs(memoryStream);

                            memoryStream.Position = 0;

                            returnedMemoryStream = memoryStream;
                        }
                    }
                }

                return File.ReadAllBytes(@"" + fileName);
            }

            catch (Exception)
            {
                throw;
            }
            finally
            {
                File.Delete(fileName);
                File.Delete(templateFileName);
            }
        }


        public byte[] DownloadDashboardFourImagemPura(List<GraficoImagemEvolutivoProcedure> grafico, string tituloPlanilha, List<TraducaoComponente> listaTraducaoComponente)
        {
            string fileName = "";
            string templateFileName = "";
            try
            {
                string strDownloadNome;

                strDownloadNome = "Modelo_Exportação_Gráficos_Padrao_Four";

                string strPath = AdjustPath(HttpContext.Current.Server.MapPath("~")) + @"Arquivos\ExcelExportacao\workspace\";
                fileName = strPath + strDownloadNome + DateTime.Now.ToString("yyyy - MM - dd HH':'mm':'ss':'fff").Replace(" - ", "").Replace(":", "").Replace(" ", "") + ".xlsx";

                templateFileName = AdjustPath(HttpContext.Current.Server.MapPath("~")) + @"Arquivos\ExcelExportacao\Modelos\" + strDownloadNome + ".xlsx";

                GetNewFileMemory(ref templateFileName);

                if (File.Exists(fileName))
                    File.Delete(fileName);

                MemoryStream returnedMemoryStream = null;

                using (var file = new FileStream(fileName, FileMode.CreateNew))
                {
                    using (var temp = new FileStream(templateFileName, FileMode.Open))
                    {
                        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                        ExcelPackage package = new ExcelPackage(file, temp);

                        ExcelWorksheet wworksheet = package.Workbook.Worksheets["Data"];

                        wworksheet.View.ShowHeaders = true;
                        wworksheet.View.ShowGridLines = false;
                        //Titulo da planilha. Deixar linha 3 vazias formatadas no modelo excel.
                        wworksheet.Cells[3, 2].Value = tituloPlanilha;

                        int currentRow = 6;
                        int colNameTable = 3;
                        int lastColName = 0;


                        if (grafico != null)
                        {

                            // Preenche cabeçalho com os periodos:
                            foreach (var item in grafico.Select(a => a.DescMarca).Distinct().ToList())
                            {
                                wworksheet.Cells[currentRow, colNameTable].Value = item;
                                colNameTable++;
                                wworksheet.Cells[currentRow, colNameTable].Value = "Base";
                                colNameTable++;
                            }

                            colNameTable--;
                            lastColName = colNameTable;

                            //Formata cabeçalho
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(3, 6) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Font.Size = 11;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(3, 6) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Font.Name = "Arial Nova Cond";
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(3, 6) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Font.Bold = true;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(3, 6) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(3, 6) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(3, 6) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(3, 6) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(3, 6) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(3, 6) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                            colNameTable = 2;
                            currentRow++;

                            ////Preenche lateral
                            foreach (var item in grafico.Select(a => a.DescOnda).Distinct().ToList())
                            {
                                wworksheet.Cells[currentRow, colNameTable].Value = item;
                                currentRow++;
                            }

                            currentRow--;
                            var lastRow = currentRow;
                            colNameTable = 3;
                            currentRow = 7;

                            foreach (var item in grafico.Select(a => a.DescOnda).Distinct().ToList())
                            {
                                colNameTable = 3;
                                var itemLinha = grafico.Where(a => a.DescOnda == item).ToList();

                                // MONTA COLUNAS DADOS
                                foreach (var dadoLinha in itemLinha)
                                {
                                    wworksheet.Cells[currentRow, colNameTable].Value = dadoLinha.Perc;
                                    if (dadoLinha.TesteSig == "MENOR")
                                        wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Red);
                                    if (dadoLinha.TesteSig == "MAIOR")
                                        wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Blue);
                                    colNameTable++;

                                    wworksheet.Cells[currentRow, colNameTable].Value = dadoLinha.BaseAbs;
                                    colNameTable++;
                                }

                                currentRow++;
                            }

                            //currentRow = 7;

                            //lastRow++;
                            //currentRow = lastRow;
                            colNameTable--;
                            //lastColName = colNameTable;

                            //Formata dados
                            //wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 7) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Numberformat.Format = "0.00";
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 7) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Font.Size = 11;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 7) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Font.Name = "Arial Nova Cond";
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 7) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Font.Bold = true;

                            // Preenche lateral
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, currentRow) + ":" + ExcelUtil.GetCelulaByColumnIndex(3, currentRow)].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, currentRow) + ":" + ExcelUtil.GetCelulaByColumnIndex(3, currentRow)].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, currentRow) + ":" + ExcelUtil.GetCelulaByColumnIndex(3, currentRow)].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, currentRow) + ":" + ExcelUtil.GetCelulaByColumnIndex(3, currentRow)].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, currentRow) + ":" + ExcelUtil.GetCelulaByColumnIndex(3, currentRow)].Style.Border.Top.Style = ExcelBorderStyle.Thin;

                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 7) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 7) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 7) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 7) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 7) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                        }
                        wworksheet.Select();

                        using (MemoryStream memoryStream = new MemoryStream())
                        {
                            package.SaveAs(memoryStream);

                            memoryStream.Position = 0;

                            returnedMemoryStream = memoryStream;
                        }
                    }
                }

                return File.ReadAllBytes(@"" + fileName);
            }

            catch (Exception)
            {
                throw;
            }
            finally
            {
                File.Delete(fileName);
                File.Delete(templateFileName);
            }
        }

        #endregion


        #region DownloadDashboard Comuns
        public byte[] DownloadDashboardFiveComparativoMarcas(GraficoImagemPosicionamentoFullLoad grafico, string tituloPlanilha, List<TraducaoComponente> listaTraducaoComponente)
        {
            string fileName = "";
            string templateFileName = "";
            try
            {
                string strDownloadNome;

                strDownloadNome = "Modelo_Exportação_Gráficos_Padrao";

                string strPath = AdjustPath(HttpContext.Current.Server.MapPath("~")) + @"Arquivos\ExcelExportacao\workspace\";
                fileName = strPath + strDownloadNome + DateTime.Now.ToString("yyyy - MM - dd HH':'mm':'ss':'fff").Replace(" - ", "").Replace(":", "").Replace(" ", "") + ".xlsx";

                templateFileName = AdjustPath(HttpContext.Current.Server.MapPath("~")) + @"Arquivos\ExcelExportacao\Modelos\" + strDownloadNome + ".xlsx";

                GetNewFileMemory(ref templateFileName);

                if (File.Exists(fileName))
                    File.Delete(fileName);

                MemoryStream returnedMemoryStream = null;

                using (var file = new FileStream(fileName, FileMode.CreateNew))
                {
                    using (var temp = new FileStream(templateFileName, FileMode.Open))
                    {
                        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                        ExcelPackage package = new ExcelPackage(file, temp);

                        ExcelWorksheet wworksheet = package.Workbook.Worksheets["Data"];

                        wworksheet.View.ShowHeaders = true;
                        wworksheet.View.ShowGridLines = false;
                        //Titulo da planilha. Deixar linha 3 vazias formatadas no modelo excel.
                        wworksheet.Cells[3, 2].Value = tituloPlanilha;

                        int currentRow = 6;
                        int colNameTable = 3;
                        int lastColName = 0;


                        if (grafico != null)
                        {
                            // Preenche cabeçalho com os periodos:
                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoImagemPosicionamento1.Any() ? grafico.GraficoImagemPosicionamento1.FirstOrDefault().DescMarca : "Sem dados";
                            colNameTable++;
                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoImagemPosicionamento2.Any() ? grafico.GraficoImagemPosicionamento2.FirstOrDefault().DescMarca : "Sem dados";
                            colNameTable++;

                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoImagemPosicionamento3.Any() ? grafico.GraficoImagemPosicionamento3.FirstOrDefault().DescMarca : "Sem dados";
                            colNameTable++;

                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoImagemPosicionamento4.Any() ? grafico.GraficoImagemPosicionamento4.FirstOrDefault().DescMarca : "Sem dados";
                            colNameTable++;

                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoImagemPosicionamento5.Any() ? grafico.GraficoImagemPosicionamento5.FirstOrDefault().DescMarca : "Sem dados";
                            colNameTable++;

                            colNameTable--;
                            lastColName = colNameTable;

                            //Formata cabeçalho
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(3, 6) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Font.Size = 11;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(3, 6) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Font.Name = "Arial Nova Cond";
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(3, 6) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Font.Bold = true;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(3, 6) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(3, 6) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(3, 6) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(3, 6) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(3, 6) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(3, 6) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                            colNameTable = 2;
                            currentRow++;


                            ////Preenche lateral

                            foreach (var item in grafico.GraficoImagemPosicionamento1)
                            {
                                wworksheet.Cells[currentRow, colNameTable].Value = item.Descricao;
                                currentRow++;
                            }

                            wworksheet.Cells[currentRow, colNameTable].Value = "Base";
                            currentRow++;

                            currentRow--;
                            var lastRow = currentRow;
                            colNameTable = 3;
                            currentRow = 6;


                            foreach (var item in grafico.GraficoImagemPosicionamento1)
                            {
                                currentRow++;
                                wworksheet.Cells[currentRow, colNameTable].Value = item.PercAtual;
                                if (item.PercAtual < -20 || item.PercAtual > 20)
                                {
                                    Color colour = ColorTranslator.FromHtml(item.CorSiteMarca);
                                    wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(colour);
                                }


                            }
                            currentRow++;
                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoImagemPosicionamento1.FirstOrDefault().Base;


                            currentRow = 6;
                            colNameTable++;
                            foreach (var item in grafico.GraficoImagemPosicionamento2)
                            {
                                currentRow++;
                                wworksheet.Cells[currentRow, colNameTable].Value = item.PercAtual;
                                if (item.PercAtual < -20 || item.PercAtual > 20)
                                {
                                    Color colour = ColorTranslator.FromHtml(item.CorSiteMarca);
                                    wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(colour);
                                }
                            }
                            currentRow++;
                            if (grafico.GraficoImagemPosicionamento2.FirstOrDefault() == null)
                                wworksheet.Cells[currentRow, colNameTable].Value = "0";
                            else
                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoImagemPosicionamento2.FirstOrDefault().Base;

                            currentRow = 6;
                            colNameTable++;
                            foreach (var item in grafico.GraficoImagemPosicionamento3)
                            {
                                currentRow++;
                                wworksheet.Cells[currentRow, colNameTable].Value = item.PercAtual;
                                if (item.PercAtual < -20 || item.PercAtual > 20)
                                {
                                    Color colour = ColorTranslator.FromHtml(item.CorSiteMarca);
                                    wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(colour);
                                }
                            }
                            currentRow++;
                            if (grafico.GraficoImagemPosicionamento3.FirstOrDefault() == null)
                                wworksheet.Cells[currentRow, colNameTable].Value = "0";
                            else
                                wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoImagemPosicionamento3.FirstOrDefault().Base;

                            currentRow = 6;
                            colNameTable++;
                            foreach (var item in grafico.GraficoImagemPosicionamento4)
                            {
                                currentRow++;
                                wworksheet.Cells[currentRow, colNameTable].Value = item.PercAtual;
                                if (item.PercAtual < -20 || item.PercAtual > 20)
                                {
                                    Color colour = ColorTranslator.FromHtml(item.CorSiteMarca);
                                    wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(colour);
                                }
                            }
                            currentRow++;
                            if (grafico.GraficoImagemPosicionamento4.FirstOrDefault() == null)
                                wworksheet.Cells[currentRow, colNameTable].Value = "0";
                            else
                                wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoImagemPosicionamento4.FirstOrDefault().Base;

                            currentRow = 6;
                            colNameTable++;
                            foreach (var item in grafico.GraficoImagemPosicionamento5)
                            {
                                currentRow++;
                                wworksheet.Cells[currentRow, colNameTable].Value = item.PercAtual;
                                if (item.PercAtual < -20 || item.PercAtual > 20)
                                {
                                    Color colour = ColorTranslator.FromHtml(item.CorSiteMarca);
                                    wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(colour);
                                }
                            }
                            currentRow++;
                            if (grafico.GraficoImagemPosicionamento5.FirstOrDefault() == null)
                                wworksheet.Cells[currentRow, colNameTable].Value = "0";
                            else
                                wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoImagemPosicionamento5.FirstOrDefault().Base;

                            colNameTable++;
                            currentRow = 7;

                            lastRow++;
                            currentRow = lastRow - 1;
                            colNameTable--;
                            lastColName = colNameTable;

                            //Formata dados
                            //wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 7) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Numberformat.Format = "0.00";
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 7) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Font.Size = 11;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 7) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Font.Name = "Arial Nova Cond";
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 7) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Font.Bold = true;

                            // Preenche lateral
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 11) + ":" + ExcelUtil.GetCelulaByColumnIndex(3, 11)].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 11) + ":" + ExcelUtil.GetCelulaByColumnIndex(3, 11)].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 11) + ":" + ExcelUtil.GetCelulaByColumnIndex(3, 11)].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 11) + ":" + ExcelUtil.GetCelulaByColumnIndex(3, 11)].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 11) + ":" + ExcelUtil.GetCelulaByColumnIndex(3, 11)].Style.Border.Top.Style = ExcelBorderStyle.Thin;

                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 7) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 7) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 7) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 7) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 7) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                        }
                        wworksheet.Select();

                        using (MemoryStream memoryStream = new MemoryStream())
                        {
                            package.SaveAs(memoryStream);

                            memoryStream.Position = 0;

                            returnedMemoryStream = memoryStream;
                        }
                    }
                }

                return File.ReadAllBytes(@"" + fileName);
            }

            catch (Exception)
            {
                throw;
            }
            finally
            {
                File.Delete(fileName);
                File.Delete(templateFileName);
            }
        }


        public byte[] DownloadDashboardFiveEvolutivoMarcas(GraficoImagemPosicionamentoFullLoad grafico, string tituloPlanilha, List<TraducaoComponente> listaTraducaoComponente, FiltroPadraoExcel filtro)
        {
            string fileName = "";
            string templateFileName = "";
            try
            {
                string strDownloadNome;

                strDownloadNome = "Modelo_Exportação_Gráficos_Padrao";

                string strPath = AdjustPath(HttpContext.Current.Server.MapPath("~")) + @"Arquivos\ExcelExportacao\workspace\";
                fileName = strPath + strDownloadNome + DateTime.Now.ToString("yyyy - MM - dd HH':'mm':'ss':'fff").Replace(" - ", "").Replace(":", "").Replace(" ", "") + ".xlsx";

                templateFileName = AdjustPath(HttpContext.Current.Server.MapPath("~")) + @"Arquivos\ExcelExportacao\Modelos\" + strDownloadNome + ".xlsx";

                GetNewFileMemory(ref templateFileName);

                if (File.Exists(fileName))
                    File.Delete(fileName);

                MemoryStream returnedMemoryStream = null;

                using (var file = new FileStream(fileName, FileMode.CreateNew))
                {
                    using (var temp = new FileStream(templateFileName, FileMode.Open))
                    {
                        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                        ExcelPackage package = new ExcelPackage(file, temp);

                        ExcelWorksheet wworksheet = package.Workbook.Worksheets["Data"];

                        wworksheet.View.ShowHeaders = true;
                        wworksheet.View.ShowGridLines = false;
                        //Titulo da planilha. Deixar linha 3 vazias formatadas no modelo excel.
                        wworksheet.Cells[3, 2].Value = tituloPlanilha;

                        int currentRow = 6;
                        int colNameTable = 3;
                        int lastColName = 0;


                        if (grafico != null)
                        {
                            // Preenche cabeçalho com os periodos:
                            if(filtro.Onda1.FirstOrDefault()==null)
                            { wworksheet.Cells[currentRow, colNameTable].Value = "N/A"; }
                            else
                            wworksheet.Cells[currentRow, colNameTable].Value = filtro.Onda1.FirstOrDefault().DescItem;
                            colNameTable++;

                            if (filtro.Onda2.FirstOrDefault() == null)
                            { wworksheet.Cells[currentRow, colNameTable].Value = "N/A"; }
                            else
                                wworksheet.Cells[currentRow, colNameTable].Value = filtro.Onda2.FirstOrDefault().DescItem;
                            colNameTable++;

                            if (filtro.Onda3.FirstOrDefault() == null)
                            { wworksheet.Cells[currentRow, colNameTable].Value = "N/A"; }
                            else
                                wworksheet.Cells[currentRow, colNameTable].Value = filtro.Onda3.FirstOrDefault().DescItem;
                            colNameTable++;

                            if (filtro.Onda4.FirstOrDefault() == null)
                            { wworksheet.Cells[currentRow, colNameTable].Value = "N/A"; }
                            else
                                wworksheet.Cells[currentRow, colNameTable].Value = filtro.Onda4.FirstOrDefault().DescItem;
                            colNameTable++;

                            colNameTable--;
                            lastColName = colNameTable;

                            //Formata cabeçalho
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(3, 6) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Font.Size = 11;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(3, 6) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Font.Name = "Arial Nova Cond";
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(3, 6) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Font.Bold = true;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(3, 6) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(3, 6) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(3, 6) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(3, 6) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(3, 6) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(3, 6) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                            colNameTable = 2;
                            currentRow++;


                            ////Preenche lateral
                            currentRow--;

                            if (grafico.GraficoImagemPosicionamento1.FirstOrDefault() == null)
                            {
                                wworksheet.Cells[currentRow, colNameTable].Value = "0";
                            }
                            else
                            {
                                wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoImagemPosicionamento1.FirstOrDefault().DescMarca;
                                Color colourMarca = ColorTranslator.FromHtml(grafico.GraficoImagemPosicionamento1.FirstOrDefault().CorSiteMarca);
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(colourMarca);
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Size = 14;
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Name = "Arial Nova Cond";
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Bold = true;
                            }
                            currentRow++;

                            //foreach (var item in grafico.GraficoImagemPosicionamento1)
                            //{
                            //    wworksheet.Cells[currentRow, colNameTable].Value = item.Descricao;
                            //    currentRow++;
                            //}

                            var fontTitulo = grafico.GraficoImagemPosicionamento1.Any(x => !string.IsNullOrWhiteSpace(x.Negrito));

                            foreach (var item in grafico.GraficoImagemPosicionamento1)
                            {
                                wworksheet.Cells[currentRow, colNameTable].Value = item.Descricao;

                                if (fontTitulo)
                                {
                                    if (!string.IsNullOrWhiteSpace(item.Negrito))
                                    {
                                        wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Black);
                                    }
                                    else
                                    {
                                        Color colour = ColorTranslator.FromHtml("#4F4F4F");
                                        wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(colour);

                                    }
                                }


                                currentRow++;
                            }


                            wworksheet.Cells[currentRow, colNameTable].Value = "Base";
                            currentRow++;

                            currentRow--;
                            var lastRow = currentRow;
                            colNameTable = 3;
                            currentRow = 6;


                            foreach (var item in grafico.GraficoImagemPosicionamento1)
                            {
                                currentRow++;
                                wworksheet.Cells[currentRow, colNameTable].Value = item.PercAtual;
                                if (item.PercAtual < -20 || item.PercAtual > 20)
                                {
                                    Color colour = ColorTranslator.FromHtml(item.CorSiteMarca);
                                    wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(colour);
                                }


                            }
                            currentRow++;

                            if(grafico.GraficoImagemPosicionamento1.FirstOrDefault()==null)
                            {
                                wworksheet.Cells[currentRow, colNameTable].Value = "0";
                            }
                            else
                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoImagemPosicionamento1.FirstOrDefault().Base;


                            currentRow = 6;
                            colNameTable++;
                            foreach (var item in grafico.GraficoImagemPosicionamento2)
                            {
                                currentRow++;
                                wworksheet.Cells[currentRow, colNameTable].Value = item.PercAtual;
                                if (item.PercAtual < -20 || item.PercAtual > 20)
                                {
                                    Color colour = ColorTranslator.FromHtml(item.CorSiteMarca);
                                    wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(colour);
                                }
                            }
                            currentRow++;

                            if (grafico.GraficoImagemPosicionamento2.FirstOrDefault() == null)
                            {
                                wworksheet.Cells[currentRow, colNameTable].Value = "0";
                            }
                            else
                                wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoImagemPosicionamento2.FirstOrDefault().Base;

                            currentRow = 6;
                            colNameTable++;
                            foreach (var item in grafico.GraficoImagemPosicionamento3)
                            {
                                currentRow++;
                                wworksheet.Cells[currentRow, colNameTable].Value = item.PercAtual;
                                if (item.PercAtual < -20 || item.PercAtual > 20)
                                {
                                    Color colour = ColorTranslator.FromHtml(item.CorSiteMarca);
                                    wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(colour);
                                }
                            }
                            currentRow++;

                            if (grafico.GraficoImagemPosicionamento3.FirstOrDefault() == null)
                            {
                                wworksheet.Cells[currentRow, colNameTable].Value = "0";
                            }
                            else
                                wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoImagemPosicionamento3.FirstOrDefault().Base;

                            currentRow = 6;
                            colNameTable++;
                            foreach (var item in grafico.GraficoImagemPosicionamento4)
                            {
                                currentRow++;
                                wworksheet.Cells[currentRow, colNameTable].Value = item.PercAtual;
                                if (item.PercAtual < -20 || item.PercAtual > 20)
                                {
                                    Color colour = ColorTranslator.FromHtml(item.CorSiteMarca);
                                    wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(colour);
                                }
                            }
                            currentRow++;
                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoImagemPosicionamento4.FirstOrDefault().Base;



                            colNameTable++;
                            currentRow = 7;

                            lastRow++;
                            currentRow = lastRow - 1;
                            colNameTable--;
                            lastColName = colNameTable;

                            //Formata dados
                            //wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 7) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Numberformat.Format = "0.00";
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 7) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Font.Size = 11;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 7) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Font.Name = "Arial Nova Cond";
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 7) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Font.Bold = true;

                            // Preenche lateral
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 11) + ":" + ExcelUtil.GetCelulaByColumnIndex(3, 11)].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 11) + ":" + ExcelUtil.GetCelulaByColumnIndex(3, 11)].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 11) + ":" + ExcelUtil.GetCelulaByColumnIndex(3, 11)].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 11) + ":" + ExcelUtil.GetCelulaByColumnIndex(3, 11)].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 11) + ":" + ExcelUtil.GetCelulaByColumnIndex(3, 11)].Style.Border.Top.Style = ExcelBorderStyle.Thin;

                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 7) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 7) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 7) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 7) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 7) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                        }
                        wworksheet.Select();

                        using (MemoryStream memoryStream = new MemoryStream())
                        {
                            package.SaveAs(memoryStream);

                            memoryStream.Position = 0;

                            returnedMemoryStream = memoryStream;
                        }
                    }
                }

                return File.ReadAllBytes(@"" + fileName);
            }

            catch (Exception)
            {
                throw;
            }
            finally
            {
                File.Delete(fileName);
                File.Delete(templateFileName);
            }
        }


        public byte[] DownloadDashboardFiveComparativoDuploMarcas(GraficoImagemPosicionamentoFullLoad grafico, string tituloPlanilha, List<TraducaoComponente> listaTraducaoComponente, FiltroPadraoExcelDuplo filtro)
        {
            string fileName = "";
            string templateFileName = "";
            try
            {
                string strDownloadNome;

                strDownloadNome = "Modelo_Exportação_Gráficos_Padrao";

                string strPath = AdjustPath(HttpContext.Current.Server.MapPath("~")) + @"Arquivos\ExcelExportacao\workspace\";
                fileName = strPath + strDownloadNome + DateTime.Now.ToString("yyyy - MM - dd HH':'mm':'ss':'fff").Replace(" - ", "").Replace(":", "").Replace(" ", "") + ".xlsx";

                templateFileName = AdjustPath(HttpContext.Current.Server.MapPath("~")) + @"Arquivos\ExcelExportacao\Modelos\" + strDownloadNome + ".xlsx";

                GetNewFileMemory(ref templateFileName);

                if (File.Exists(fileName))
                    File.Delete(fileName);

                MemoryStream returnedMemoryStream = null;

                using (var file = new FileStream(fileName, FileMode.CreateNew))
                {
                    using (var temp = new FileStream(templateFileName, FileMode.Open))
                    {
                        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                        ExcelPackage package = new ExcelPackage(file, temp);

                        ExcelWorksheet wworksheet = package.Workbook.Worksheets["Data"];

                        wworksheet.View.ShowHeaders = true;
                        wworksheet.View.ShowGridLines = false;
                        //Titulo da planilha. Deixar linha 3 vazias formatadas no modelo excel.
                        wworksheet.Cells[3, 2].Value = tituloPlanilha;

                        int currentRow = 6;
                        int colNameTable = 3;
                        int lastColName = 0;


                        if (grafico != null)
                        {
                            // Preenche cabeçalho com os periodos:
                            wworksheet.Cells[currentRow, colNameTable].Value = filtro.MarcaDuploColuna1.DescItem + String.Concat(" - ", filtro.OndaDuploColuna1 != null ? filtro.OndaDuploColuna1.DescItem : "N/A");
                            colNameTable++;

                            wworksheet.Cells[currentRow, colNameTable].Value = filtro.MarcaDuploColuna1.DescItem + String.Concat(" - ", filtro.OndaDuploColuna1_2 != null ? filtro.OndaDuploColuna1_2.DescItem : "N/A");
                            colNameTable++;
                            wworksheet.Cells[currentRow, colNameTable].Value = "";
                            wworksheet.Column(colNameTable).Width = 5;
                            colNameTable++;

                            wworksheet.Cells[currentRow, colNameTable].Value = filtro.MarcaDuploColuna2.DescItem + String.Concat(" - ", filtro.OndaDuploColuna2 != null ? filtro.OndaDuploColuna2.DescItem : "N/A");
                            colNameTable++;
                            wworksheet.Cells[currentRow, colNameTable].Value = filtro.MarcaDuploColuna2.DescItem + String.Concat(" - ", filtro.OndaDuploColuna2_2 != null ? filtro.OndaDuploColuna2_2.DescItem : "N/A");
                            colNameTable++;
                            wworksheet.Cells[currentRow, colNameTable].Value = "";
                            wworksheet.Column(colNameTable).Width = 5;
                            colNameTable++;

                            wworksheet.Cells[currentRow, colNameTable].Value = filtro.MarcaDuploColuna3.DescItem + String.Concat(" - ", filtro.OndaDuploColuna3 != null ? filtro.OndaDuploColuna3.DescItem : "N/A");
                            colNameTable++;
                            wworksheet.Cells[currentRow, colNameTable].Value = filtro.MarcaDuploColuna3.DescItem + String.Concat(" - ", filtro.OndaDuploColuna3_2 != null ? filtro.OndaDuploColuna3_2.DescItem : "N/A");
                            colNameTable++;
                            wworksheet.Cells[currentRow, colNameTable].Value = "";
                            wworksheet.Column(colNameTable).Width = 5;
                            colNameTable++;

                            wworksheet.Cells[currentRow, colNameTable].Value = filtro.MarcaDuploColuna4.DescItem + String.Concat(" - ", filtro.OndaDuploColuna4 != null ? filtro.OndaDuploColuna4.DescItem : "N/A");
                            colNameTable++;
                            wworksheet.Cells[currentRow, colNameTable].Value = filtro.MarcaDuploColuna4.DescItem + String.Concat(" - ", filtro.OndaDuploColuna4_2 != null ? filtro.OndaDuploColuna4_2.DescItem : "N/A");
                            colNameTable++;
                            wworksheet.Cells[currentRow, colNameTable].Value = "";
                            wworksheet.Column(colNameTable).Width = 5;
                            colNameTable++;

                            wworksheet.Cells[currentRow, colNameTable].Value = filtro.MarcaDuploColuna5.DescItem + String.Concat(" - ", filtro.OndaDuploColuna5 != null ? filtro.OndaDuploColuna5.DescItem : "N/A");
                            colNameTable++;
                            wworksheet.Cells[currentRow, colNameTable].Value = filtro.MarcaDuploColuna5.DescItem + String.Concat(" - ", filtro.OndaDuploColuna5_2 != null ? filtro.OndaDuploColuna5_2.DescItem : "");
                            colNameTable++;
                            wworksheet.Cells[currentRow, colNameTable].Value = "";
                            wworksheet.Column(colNameTable).Width = 5;
                            colNameTable++;

                            colNameTable--;
                            lastColName = colNameTable;

                            //Formata cabeçalho
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(3, 6) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Font.Size = 11;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(3, 6) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Font.Name = "Arial Nova Cond";
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(3, 6) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Font.Bold = true;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(3, 6) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(3, 6) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(3, 6) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(3, 6) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(3, 6) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(3, 6) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                            colNameTable = 2;
                            currentRow++;


                            ////Preenche lateral
                            ///

                            var fontTitulo = grafico.GraficoImagemPosicionamento1.Any(x => !string.IsNullOrWhiteSpace(x.Negrito));

                            foreach (var item in grafico.GraficoImagemPosicionamento1)
                            {
                                wworksheet.Cells[currentRow, colNameTable].Value = item.Descricao;

                                if (fontTitulo)
                                {
                                    if (!string.IsNullOrWhiteSpace(item.Negrito))
                                    {
                                        wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Black);
                                    }
                                    else
                                    {
                                        Color colour = ColorTranslator.FromHtml("#4F4F4F");
                                        wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(colour);

                                    }
                                }


                                currentRow++;
                            }

                            wworksheet.Cells[currentRow, colNameTable].Value = "Base";
                            currentRow++;

                            currentRow--;
                            var lastRow = currentRow;
                            colNameTable = 3;
                            currentRow = 6;


                            foreach (var item in grafico.GraficoImagemPosicionamento1)
                            {
                                currentRow++;
                                wworksheet.Cells[currentRow, colNameTable].Value = item.PercAtual;
                                if (item.PercAtual < -20 || item.PercAtual > 20)
                                {
                                    Color colour = ColorTranslator.FromHtml(item.CorSiteMarca);
                                    wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(colour);
                                }


                            }
                            currentRow++;
                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoImagemPosicionamento1.FirstOrDefault().Base;


                            currentRow = 6;
                            colNameTable++;
                            foreach (var item in grafico.GraficoImagemPosicionamento2)
                            {
                                currentRow++;
                                wworksheet.Cells[currentRow, colNameTable].Value = item.PercAtual;
                                if (item.PercAtual < -20 || item.PercAtual > 20)
                                {
                                    Color colour = ColorTranslator.FromHtml(item.CorSiteMarca);
                                    wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(colour);
                                }
                            }
                            currentRow++;
                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoImagemPosicionamento2.FirstOrDefault().Base;

                            // LINHA BRANCA
                            currentRow = 6;
                            colNameTable++;
                            foreach (var item in grafico.GraficoImagemPosicionamento2)
                            {
                                currentRow++;
                                wworksheet.Cells[currentRow, colNameTable].Value = "";
                            }
                            currentRow++;
                            wworksheet.Cells[currentRow, colNameTable].Value = "";


                            #region Linha 2
                            currentRow = 6;
                            colNameTable++;
                            foreach (var item in grafico.GraficoImagemPosicionamento3)
                            {
                                currentRow++;
                                wworksheet.Cells[currentRow, colNameTable].Value = item.PercAtual;
                                if (item.PercAtual < -20 || item.PercAtual > 20)
                                {
                                    Color colour = ColorTranslator.FromHtml(item.CorSiteMarca);
                                    wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(colour);
                                }


                            }
                            currentRow++;
                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoImagemPosicionamento3.FirstOrDefault().Base;


                            currentRow = 6;
                            colNameTable++;
                            foreach (var item in grafico.GraficoImagemPosicionamento4)
                            {
                                currentRow++;
                                wworksheet.Cells[currentRow, colNameTable].Value = item.PercAtual;
                                if (item.PercAtual < -20 || item.PercAtual > 20)
                                {
                                    Color colour = ColorTranslator.FromHtml(item.CorSiteMarca);
                                    wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(colour);
                                }
                            }
                            currentRow++;
                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoImagemPosicionamento4.FirstOrDefault().Base;

                            // LINHA BRANCA
                            currentRow = 6;
                            colNameTable++;
                            foreach (var item in grafico.GraficoImagemPosicionamento4)
                            {
                                currentRow++;
                                wworksheet.Cells[currentRow, colNameTable].Value = "";
                            }
                            currentRow++;
                            wworksheet.Cells[currentRow, colNameTable].Value = "";

                            #endregion


                            #region Linha 3
                            currentRow = 6;
                            colNameTable++;
                            foreach (var item in grafico.GraficoImagemPosicionamento5)
                            {
                                currentRow++;
                                wworksheet.Cells[currentRow, colNameTable].Value = item.PercAtual;
                                if (item.PercAtual < -20 || item.PercAtual > 20)
                                {
                                    Color colour = ColorTranslator.FromHtml(item.CorSiteMarca);
                                    wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(colour);
                                }


                            }
                            currentRow++;
                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoImagemPosicionamento5.FirstOrDefault().Base;


                            currentRow = 6;
                            colNameTable++;
                            foreach (var item in grafico.GraficoImagemPosicionamento6)
                            {
                                currentRow++;
                                wworksheet.Cells[currentRow, colNameTable].Value = item.PercAtual;
                                if (item.PercAtual < -20 || item.PercAtual > 20)
                                {
                                    Color colour = ColorTranslator.FromHtml(item.CorSiteMarca);
                                    wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(colour);
                                }
                            }
                            currentRow++;
                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoImagemPosicionamento6.FirstOrDefault().Base;

                            // LINHA BRANCA
                            currentRow = 6;
                            colNameTable++;
                            foreach (var item in grafico.GraficoImagemPosicionamento6)
                            {
                                currentRow++;
                                wworksheet.Cells[currentRow, colNameTable].Value = "";
                            }
                            currentRow++;
                            wworksheet.Cells[currentRow, colNameTable].Value = "";

                            #endregion


                            #region Linha 4
                            currentRow = 6;
                            colNameTable++;
                            foreach (var item in grafico.GraficoImagemPosicionamento7)
                            {
                                currentRow++;
                                wworksheet.Cells[currentRow, colNameTable].Value = item.PercAtual;
                                if (item.PercAtual < -20 || item.PercAtual > 20)
                                {
                                    Color colour = ColorTranslator.FromHtml(item.CorSiteMarca);
                                    wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(colour);
                                }


                            }
                            currentRow++;
                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoImagemPosicionamento7.FirstOrDefault().Base;


                            currentRow = 6;
                            colNameTable++;
                            foreach (var item in grafico.GraficoImagemPosicionamento8)
                            {
                                currentRow++;
                                wworksheet.Cells[currentRow, colNameTable].Value = item.PercAtual;
                                if (item.PercAtual < -20 || item.PercAtual > 20)
                                {
                                    Color colour = ColorTranslator.FromHtml(item.CorSiteMarca);
                                    wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(colour);
                                }
                            }
                            currentRow++;
                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoImagemPosicionamento8.FirstOrDefault().Base;

                            // LINHA BRANCA
                            currentRow = 6;
                            colNameTable++;
                            foreach (var item in grafico.GraficoImagemPosicionamento8)
                            {
                                currentRow++;
                                wworksheet.Cells[currentRow, colNameTable].Value = "";
                            }
                            currentRow++;
                            wworksheet.Cells[currentRow, colNameTable].Value = "";

                            #endregion


                            #region Linha 5
                            currentRow = 6;
                            colNameTable++;
                            foreach (var item in grafico.GraficoImagemPosicionamento9)
                            {
                                currentRow++;
                                wworksheet.Cells[currentRow, colNameTable].Value = item.PercAtual;
                                if (item.PercAtual < -20 || item.PercAtual > 20)
                                {
                                    Color colour = ColorTranslator.FromHtml(item.CorSiteMarca);
                                    wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(colour);
                                }


                            }
                            currentRow++;
                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoImagemPosicionamento9.FirstOrDefault().Base;


                            currentRow = 6;
                            colNameTable++;
                            foreach (var item in grafico.GraficoImagemPosicionamento10)
                            {
                                currentRow++;
                                wworksheet.Cells[currentRow, colNameTable].Value = item.PercAtual;
                                if (item.PercAtual < -20 || item.PercAtual > 20)
                                {
                                    Color colour = ColorTranslator.FromHtml(item.CorSiteMarca);
                                    wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(colour);
                                }
                            }
                            currentRow++;
                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoImagemPosicionamento10.FirstOrDefault().Base;

                            // LINHA BRANCA
                            currentRow = 6;
                            colNameTable++;
                            foreach (var item in grafico.GraficoImagemPosicionamento10)
                            {
                                currentRow++;
                                wworksheet.Cells[currentRow, colNameTable].Value = "";
                            }
                            currentRow++;
                            wworksheet.Cells[currentRow, colNameTable].Value = "";

                            #endregion


                            colNameTable++;
                            currentRow = 7;

                            lastRow++;
                            currentRow = lastRow - 1;
                            colNameTable--;
                            lastColName = colNameTable;

                            //Formata dados
                            //wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 7) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Numberformat.Format = "0.00";
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 7) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Font.Size = 11;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 7) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Font.Name = "Arial Nova Cond";
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 7) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Font.Bold = true;

                            // Preenche lateral
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 11) + ":" + ExcelUtil.GetCelulaByColumnIndex(3, 11)].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 11) + ":" + ExcelUtil.GetCelulaByColumnIndex(3, 11)].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 11) + ":" + ExcelUtil.GetCelulaByColumnIndex(3, 11)].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 11) + ":" + ExcelUtil.GetCelulaByColumnIndex(3, 11)].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 11) + ":" + ExcelUtil.GetCelulaByColumnIndex(3, 11)].Style.Border.Top.Style = ExcelBorderStyle.Thin;

                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 7) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 7) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 7) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 7) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 7) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                        }
                        wworksheet.Select();

                        using (MemoryStream memoryStream = new MemoryStream())
                        {
                            package.SaveAs(memoryStream);

                            memoryStream.Position = 0;

                            returnedMemoryStream = memoryStream;
                        }
                    }
                }

                return File.ReadAllBytes(@"" + fileName);
            }

            catch (Exception)
            {
                throw;
            }
            finally
            {
                File.Delete(fileName);
                File.Delete(templateFileName);
            }
        }

        public byte[] DownloadDashboardFiveComparativoDuploMarcasNew(GraficoImagemPosicionamentoFullLoad grafico, string tituloPlanilha, List<TraducaoComponente> listaTraducaoComponente, FiltroPadraoExcelDuplo filtro)
        {
            string fileName = "";
            string templateFileName = "";
            try
            {
                string strDownloadNome;

                strDownloadNome = "Modelo_Exportação_Gráficos_Padrao";

                string strPath = AdjustPath(HttpContext.Current.Server.MapPath("~")) + @"Arquivos\ExcelExportacao\workspace\";
                fileName = strPath + strDownloadNome + DateTime.Now.ToString("yyyy - MM - dd HH':'mm':'ss':'fff").Replace(" - ", "").Replace(":", "").Replace(" ", "") + ".xlsx";

                templateFileName = AdjustPath(HttpContext.Current.Server.MapPath("~")) + @"Arquivos\ExcelExportacao\Modelos\" + strDownloadNome + ".xlsx";

                GetNewFileMemory(ref templateFileName);

                if (File.Exists(fileName))
                    File.Delete(fileName);

                MemoryStream returnedMemoryStream = null;

                using (var file = new FileStream(fileName, FileMode.CreateNew))
                {
                    using (var temp = new FileStream(templateFileName, FileMode.Open))
                    {
                        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                        ExcelPackage package = new ExcelPackage(file, temp);

                        ExcelWorksheet wworksheet = package.Workbook.Worksheets["Data"];

                        wworksheet.View.ShowHeaders = true;
                        wworksheet.View.ShowGridLines = false;
                        //Titulo da planilha. Deixar linha 3 vazias formatadas no modelo excel.
                        wworksheet.Cells[3, 2].Value = tituloPlanilha;

                        int currentRow = 6;
                        int colNameTable = 3;
                        int lastColName = 0;


                        if (grafico != null)
                        {
                            // Preenche cabeçalho com os periodos:
                            wworksheet.Cells[currentRow, colNameTable].Value = filtro.MarcaDuploColuna1.DescItem + String.Concat(" - ", filtro.OndaDuploColuna1.DescItem);
                            colNameTable++;

                            wworksheet.Cells[currentRow, colNameTable].Value = filtro.MarcaDuploColuna1.DescItem + String.Concat(" - ", filtro.OndaDuploColuna1_2.DescItem);
                            colNameTable++;

                            wworksheet.Cells[currentRow, colNameTable].Value = filtro.MarcaDuploColuna1.DescItem + String.Concat(" - ", filtro.OndaDuploColuna1_3.DescItem);
                            colNameTable++;
                            wworksheet.Cells[currentRow, colNameTable].Value = "";
                            wworksheet.Column(colNameTable).Width = 5;
                            colNameTable++;

                            wworksheet.Cells[currentRow, colNameTable].Value = filtro.MarcaDuploColuna2.DescItem + String.Concat(" - ", filtro.OndaDuploColuna2.DescItem);
                            colNameTable++;
                            wworksheet.Cells[currentRow, colNameTable].Value = filtro.MarcaDuploColuna2.DescItem + String.Concat(" - ", filtro.OndaDuploColuna2_2.DescItem);
                            colNameTable++;
                            wworksheet.Cells[currentRow, colNameTable].Value = filtro.MarcaDuploColuna2.DescItem + String.Concat(" - ", filtro.OndaDuploColuna2_3.DescItem);
                            colNameTable++;
                            wworksheet.Cells[currentRow, colNameTable].Value = "";
                            wworksheet.Column(colNameTable).Width = 5;
                            colNameTable++;

                            wworksheet.Cells[currentRow, colNameTable].Value = filtro.MarcaDuploColuna3.DescItem + String.Concat(" - ", filtro.OndaDuploColuna3.DescItem);
                            colNameTable++;
                            wworksheet.Cells[currentRow, colNameTable].Value = filtro.MarcaDuploColuna3.DescItem + String.Concat(" - ", filtro.OndaDuploColuna3_2.DescItem);
                            colNameTable++;
                            wworksheet.Cells[currentRow, colNameTable].Value = filtro.MarcaDuploColuna3.DescItem + String.Concat(" - ", filtro.OndaDuploColuna3_3.DescItem);
                            colNameTable++;
                            wworksheet.Cells[currentRow, colNameTable].Value = "";
                            wworksheet.Column(colNameTable).Width = 5;
                            colNameTable++;

                            wworksheet.Cells[currentRow, colNameTable].Value = filtro.MarcaDuploColuna4.DescItem + String.Concat(" - ", filtro.OndaDuploColuna4.DescItem);
                            colNameTable++;
                            wworksheet.Cells[currentRow, colNameTable].Value = filtro.MarcaDuploColuna4.DescItem + String.Concat(" - ", filtro.OndaDuploColuna4_2.DescItem);
                            colNameTable++;
                            wworksheet.Cells[currentRow, colNameTable].Value = filtro.MarcaDuploColuna4.DescItem + String.Concat(" - ", filtro.OndaDuploColuna4_3.DescItem);
                            colNameTable++;
                            wworksheet.Cells[currentRow, colNameTable].Value = "";
                            wworksheet.Column(colNameTable).Width = 5;
                            colNameTable++;

                            //wworksheet.Cells[currentRow, colNameTable].Value = filtro.MarcaDuploColuna5.DescItem + String.Concat(" - ", filtro.OndaDuploColuna5.DescItem);
                            //colNameTable++;
                            //wworksheet.Cells[currentRow, colNameTable].Value = filtro.MarcaDuploColuna5.DescItem + String.Concat(" - ", filtro.OndaDuploColuna5_2.DescItem);
                            //colNameTable++;
                            //wworksheet.Cells[currentRow, colNameTable].Value = filtro.MarcaDuploColuna5.DescItem + String.Concat(" - ", filtro.OndaDuploColuna5_3.DescItem);
                            //colNameTable++;
                            //wworksheet.Cells[currentRow, colNameTable].Value = "";
                            //wworksheet.Column(colNameTable).Width = 5;
                            //colNameTable++;

                            colNameTable--;
                            lastColName = colNameTable;

                            //Formata cabeçalho
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(3, 6) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Font.Size = 11;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(3, 6) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Font.Name = "Arial Nova Cond";
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(3, 6) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Font.Bold = true;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(3, 6) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(3, 6) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(3, 6) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(3, 6) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(3, 6) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(3, 6) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                            colNameTable = 2;
                            currentRow++;


                            ////Preenche lateral
                            ///

                            var fontTitulo = grafico.GraficoImagemPosicionamento1.Any(x => !string.IsNullOrWhiteSpace(x.Negrito));

                            foreach (var item in grafico.GraficoImagemPosicionamento1)
                            {
                                wworksheet.Cells[currentRow, colNameTable].Value = item.Descricao;

                                if (fontTitulo)
                                {
                                    if (!string.IsNullOrWhiteSpace(item.Negrito))
                                    {
                                        wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Black);
                                    }
                                    else
                                    {
                                        Color colour = ColorTranslator.FromHtml("#4F4F4F");
                                        wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(colour);

                                    }
                                }


                                currentRow++;
                            }

                            wworksheet.Cells[currentRow, colNameTable].Value = "Base";
                            currentRow++;

                            currentRow--;
                            var lastRow = currentRow;
                            colNameTable = 3;
                            currentRow = 6;


                            foreach (var item in grafico.GraficoImagemPosicionamento1)
                            {
                                currentRow++;
                                wworksheet.Cells[currentRow, colNameTable].Value = item.PercAtual;
                                if (item.PercAtual < -20 || item.PercAtual > 20)
                                {
                                    Color colour = ColorTranslator.FromHtml(item.CorSiteMarca);
                                    wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(colour);
                                }


                            }
                            currentRow++;

                            if (grafico.GraficoImagemPosicionamento1.FirstOrDefault() == null)
                            {
                                wworksheet.Cells[currentRow, colNameTable].Value = "0";
                            }
                            else
                                wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoImagemPosicionamento1.FirstOrDefault().Base;


                            currentRow = 6;
                            colNameTable++;
                            foreach (var item in grafico.GraficoImagemPosicionamento2)
                            {
                                currentRow++;
                                wworksheet.Cells[currentRow, colNameTable].Value = item.PercAtual;
                                if (item.PercAtual < -20 || item.PercAtual > 20)
                                {
                                    Color colour = ColorTranslator.FromHtml(item.CorSiteMarca);
                                    wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(colour);
                                }
                            }
                            currentRow++;

                            if (grafico.GraficoImagemPosicionamento2.FirstOrDefault() == null)
                            {
                                wworksheet.Cells[currentRow, colNameTable].Value = "0";
                            }
                            else
                                wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoImagemPosicionamento2.FirstOrDefault().Base;



                            currentRow = 6;
                            colNameTable++;
                            foreach (var item in grafico.GraficoImagemPosicionamento1_3)
                            {
                                currentRow++;
                                wworksheet.Cells[currentRow, colNameTable].Value = item.PercAtual;
                                if (item.PercAtual < -20 || item.PercAtual > 20)
                                {
                                    Color colour = ColorTranslator.FromHtml(item.CorSiteMarca);
                                    wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(colour);
                                }
                            }
                            currentRow++;

                            if (grafico.GraficoImagemPosicionamento1_3.FirstOrDefault() == null)
                            {
                                wworksheet.Cells[currentRow, colNameTable].Value = "0";
                            }
                            else
                                wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoImagemPosicionamento1_3.FirstOrDefault().Base;




                            // LINHA BRANCA
                            currentRow = 6;
                            colNameTable++;
                            foreach (var item in grafico.GraficoImagemPosicionamento2)
                            {
                                currentRow++;
                                wworksheet.Cells[currentRow, colNameTable].Value = "";
                            }
                            currentRow++;
                            wworksheet.Cells[currentRow, colNameTable].Value = "";


                            #region Linha 2
                            currentRow = 6;
                            colNameTable++;
                            foreach (var item in grafico.GraficoImagemPosicionamento3)
                            {
                                currentRow++;
                                wworksheet.Cells[currentRow, colNameTable].Value = item.PercAtual;
                                if (item.PercAtual < -20 || item.PercAtual > 20)
                                {
                                    Color colour = ColorTranslator.FromHtml(item.CorSiteMarca);
                                    wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(colour);
                                }


                            }
                            currentRow++;

                            if (grafico.GraficoImagemPosicionamento3.FirstOrDefault() == null)
                            {
                                wworksheet.Cells[currentRow, colNameTable].Value = "0";
                            }
                            else
                                wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoImagemPosicionamento3.FirstOrDefault().Base;


                            currentRow = 6;
                            colNameTable++;
                            foreach (var item in grafico.GraficoImagemPosicionamento4)
                            {
                                currentRow++;
                                wworksheet.Cells[currentRow, colNameTable].Value = item.PercAtual;
                                if (item.PercAtual < -20 || item.PercAtual > 20)
                                {
                                    Color colour = ColorTranslator.FromHtml(item.CorSiteMarca);
                                    wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(colour);
                                }
                            }
                            currentRow++;

                            if (grafico.GraficoImagemPosicionamento4.FirstOrDefault() == null)
                            {
                                wworksheet.Cells[currentRow, colNameTable].Value = "0";
                            }
                            else
                                wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoImagemPosicionamento4.FirstOrDefault().Base;


                            currentRow = 6;
                            colNameTable++;
                            foreach (var item in grafico.GraficoImagemPosicionamento3_3)
                            {
                                currentRow++;
                                wworksheet.Cells[currentRow, colNameTable].Value = item.PercAtual;
                                if (item.PercAtual < -20 || item.PercAtual > 20)
                                {
                                    Color colour = ColorTranslator.FromHtml(item.CorSiteMarca);
                                    wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(colour);
                                }
                            }
                            currentRow++;

                            if (grafico.GraficoImagemPosicionamento3_3.FirstOrDefault() == null)
                            {
                                wworksheet.Cells[currentRow, colNameTable].Value = "0";
                            }
                            else
                                wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoImagemPosicionamento3_3.FirstOrDefault().Base;

                            // LINHA BRANCA
                            currentRow = 6;
                            colNameTable++;
                            foreach (var item in grafico.GraficoImagemPosicionamento4)
                            {
                                currentRow++;
                                wworksheet.Cells[currentRow, colNameTable].Value = "";
                            }
                            currentRow++;
                            wworksheet.Cells[currentRow, colNameTable].Value = "";

                            #endregion


                            #region Linha 3
                            currentRow = 6;
                            colNameTable++;
                            foreach (var item in grafico.GraficoImagemPosicionamento5)
                            {
                                currentRow++;
                                wworksheet.Cells[currentRow, colNameTable].Value = item.PercAtual;
                                if (item.PercAtual < -20 || item.PercAtual > 20)
                                {
                                    Color colour = ColorTranslator.FromHtml(item.CorSiteMarca);
                                    wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(colour);
                                }


                            }
                            currentRow++;

                            if (grafico.GraficoImagemPosicionamento5.FirstOrDefault() == null)
                            {
                                wworksheet.Cells[currentRow, colNameTable].Value = "0";
                            }
                            else
                                wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoImagemPosicionamento5.FirstOrDefault().Base;


                            currentRow = 6;
                            colNameTable++;
                            foreach (var item in grafico.GraficoImagemPosicionamento6)
                            {
                                currentRow++;
                                wworksheet.Cells[currentRow, colNameTable].Value = item.PercAtual;
                                if (item.PercAtual < -20 || item.PercAtual > 20)
                                {
                                    Color colour = ColorTranslator.FromHtml(item.CorSiteMarca);
                                    wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(colour);
                                }
                            }
                            currentRow++;


                            if (grafico.GraficoImagemPosicionamento6.FirstOrDefault() == null)
                            {
                                wworksheet.Cells[currentRow, colNameTable].Value = "0";
                            }
                            else
                                wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoImagemPosicionamento6.FirstOrDefault().Base;

                            currentRow = 6;
                            colNameTable++;
                            foreach (var item in grafico.GraficoImagemPosicionamento5_3)
                            {
                                currentRow++;
                                wworksheet.Cells[currentRow, colNameTable].Value = item.PercAtual;
                                if (item.PercAtual < -20 || item.PercAtual > 20)
                                {
                                    Color colour = ColorTranslator.FromHtml(item.CorSiteMarca);
                                    wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(colour);
                                }
                            }
                            currentRow++;


                            if (grafico.GraficoImagemPosicionamento5_3.FirstOrDefault() == null)
                            {
                                wworksheet.Cells[currentRow, colNameTable].Value = "0";
                            }
                            else
                                wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoImagemPosicionamento5_3.FirstOrDefault().Base;

                            // LINHA BRANCA
                            currentRow = 6;
                            colNameTable++;
                            foreach (var item in grafico.GraficoImagemPosicionamento6)
                            {
                                currentRow++;
                                wworksheet.Cells[currentRow, colNameTable].Value = "";
                            }
                            currentRow++;
                            wworksheet.Cells[currentRow, colNameTable].Value = "";

                            #endregion


                            #region Linha 4
                            currentRow = 6;
                            colNameTable++;
                            foreach (var item in grafico.GraficoImagemPosicionamento7)
                            {
                                currentRow++;
                                wworksheet.Cells[currentRow, colNameTable].Value = item.PercAtual;
                                if (item.PercAtual < -20 || item.PercAtual > 20)
                                {
                                    Color colour = ColorTranslator.FromHtml(item.CorSiteMarca);
                                    wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(colour);
                                }


                            }
                            currentRow++;
                            if (grafico.GraficoImagemPosicionamento7.FirstOrDefault() == null)
                            {
                                wworksheet.Cells[currentRow, colNameTable].Value = "0";
                            }
                            else
                                wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoImagemPosicionamento7.FirstOrDefault().Base;


                            currentRow = 6;
                            colNameTable++;
                            foreach (var item in grafico.GraficoImagemPosicionamento8)
                            {
                                currentRow++;
                                wworksheet.Cells[currentRow, colNameTable].Value = item.PercAtual;
                                if (item.PercAtual < -20 || item.PercAtual > 20)
                                {
                                    Color colour = ColorTranslator.FromHtml(item.CorSiteMarca);
                                    wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(colour);
                                }
                            }
                            currentRow++;
                            if (grafico.GraficoImagemPosicionamento8.FirstOrDefault() == null)
                            {
                                wworksheet.Cells[currentRow, colNameTable].Value = "0";
                            }
                            else
                                wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoImagemPosicionamento8.FirstOrDefault().Base;

                            currentRow = 6;
                            colNameTable++;
                            foreach (var item in grafico.GraficoImagemPosicionamento7_3)
                            {
                                currentRow++;
                                wworksheet.Cells[currentRow, colNameTable].Value = item.PercAtual;
                                if (item.PercAtual < -20 || item.PercAtual > 20)
                                {
                                    Color colour = ColorTranslator.FromHtml(item.CorSiteMarca);
                                    wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(colour);
                                }
                            }
                            currentRow++;
                            if (grafico.GraficoImagemPosicionamento7_3.FirstOrDefault() == null)
                            {
                                wworksheet.Cells[currentRow, colNameTable].Value = "0";
                            }
                            else
                                wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoImagemPosicionamento7_3.FirstOrDefault().Base;

                            // LINHA BRANCA
                            currentRow = 6;
                            colNameTable++;
                            foreach (var item in grafico.GraficoImagemPosicionamento8)
                            {
                                currentRow++;
                                wworksheet.Cells[currentRow, colNameTable].Value = "";
                            }
                            currentRow++;
                            wworksheet.Cells[currentRow, colNameTable].Value = "";

                            #endregion


                            #region Linha 5
                            //currentRow = 6;
                            //colNameTable++;
                            //foreach (var item in grafico.GraficoImagemPosicionamento9)
                            //{
                            //    currentRow++;
                            //    wworksheet.Cells[currentRow, colNameTable].Value = item.PercAtual;
                            //    if (item.PercAtual < -20 || item.PercAtual > 20)
                            //    {
                            //        Color colour = ColorTranslator.FromHtml(item.CorSiteMarca);
                            //        wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(colour);
                            //    }


                            //}
                            //currentRow++;

                            //if (grafico.GraficoImagemPosicionamento9.FirstOrDefault() == null)
                            //{
                            //    wworksheet.Cells[currentRow, colNameTable].Value = "0";
                            //}
                            //else
                            //    wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoImagemPosicionamento9.FirstOrDefault().Base;


                            //currentRow = 6;
                            //colNameTable++;
                            //foreach (var item in grafico.GraficoImagemPosicionamento10)
                            //{
                            //    currentRow++;
                            //    wworksheet.Cells[currentRow, colNameTable].Value = item.PercAtual;
                            //    if (item.PercAtual < -20 || item.PercAtual > 20)
                            //    {
                            //        Color colour = ColorTranslator.FromHtml(item.CorSiteMarca);
                            //        wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(colour);
                            //    }
                            //}
                            //currentRow++;

                            //if (grafico.GraficoImagemPosicionamento10.FirstOrDefault() == null)
                            //{
                            //    wworksheet.Cells[currentRow, colNameTable].Value = "0";
                            //}
                            //else
                            //    wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoImagemPosicionamento10.FirstOrDefault().Base;


                            //currentRow = 6;
                            //colNameTable++;
                            //foreach (var item in grafico.GraficoImagemPosicionamento9_3)
                            //{
                            //    currentRow++;
                            //    wworksheet.Cells[currentRow, colNameTable].Value = item.PercAtual;
                            //    if (item.PercAtual < -20 || item.PercAtual > 20)
                            //    {
                            //        Color colour = ColorTranslator.FromHtml(item.CorSiteMarca);
                            //        wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(colour);
                            //    }
                            //}
                            //currentRow++;

                            //if (grafico.GraficoImagemPosicionamento9_3.FirstOrDefault() == null)
                            //{
                            //    wworksheet.Cells[currentRow, colNameTable].Value = "0";
                            //}
                            //else
                            //    wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoImagemPosicionamento9_3.FirstOrDefault().Base;

                            //// LINHA BRANCA
                            //currentRow = 6;
                            //colNameTable++;
                            //foreach (var item in grafico.GraficoImagemPosicionamento10)
                            //{
                            //    currentRow++;
                            //    wworksheet.Cells[currentRow, colNameTable].Value = "";
                            //}
                            //currentRow++;
                            //wworksheet.Cells[currentRow, colNameTable].Value = "";

                            #endregion


                            colNameTable++;
                            currentRow = 7;

                            lastRow++;
                            currentRow = lastRow - 1;
                            colNameTable--;
                            lastColName = colNameTable;

                            //Formata dados
                            //wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 7) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Numberformat.Format = "0.00";
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 7) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Font.Size = 11;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 7) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Font.Name = "Arial Nova Cond";
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 7) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Font.Bold = true;

                            // Preenche lateral
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 11) + ":" + ExcelUtil.GetCelulaByColumnIndex(3, 11)].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 11) + ":" + ExcelUtil.GetCelulaByColumnIndex(3, 11)].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 11) + ":" + ExcelUtil.GetCelulaByColumnIndex(3, 11)].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 11) + ":" + ExcelUtil.GetCelulaByColumnIndex(3, 11)].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 11) + ":" + ExcelUtil.GetCelulaByColumnIndex(3, 11)].Style.Border.Top.Style = ExcelBorderStyle.Thin;

                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 7) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 7) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 7) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 7) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 7) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                        }
                        wworksheet.Select();

                        using (MemoryStream memoryStream = new MemoryStream())
                        {
                            package.SaveAs(memoryStream);

                            memoryStream.Position = 0;

                            returnedMemoryStream = memoryStream;
                        }
                    }
                }

                return File.ReadAllBytes(@"" + fileName);
            }

            catch (Exception)
            {
                throw;
            }
            finally
            {
                File.Delete(fileName);
                File.Delete(templateFileName);
            }
        }
        #endregion

        public byte[] DownloadDashboardSixGraficoBVCTop10Marcas(List<GraficoBVCTop10> grafico, string tituloPlanilha, List<TraducaoComponente> listaTraducaoComponente)
        {
            string fileName = "";
            string templateFileName = "";
            try
            {
                string strDownloadNome;

                strDownloadNome = "Modelo_Exportação_Gráficos_Padrao";

                string strPath = AdjustPath(HttpContext.Current.Server.MapPath("~")) + @"Arquivos\ExcelExportacao\workspace\";
                fileName = strPath + strDownloadNome + DateTime.Now.ToString("yyyy - MM - dd HH':'mm':'ss':'fff").Replace(" - ", "").Replace(":", "").Replace(" ", "") + ".xlsx";

                templateFileName = AdjustPath(HttpContext.Current.Server.MapPath("~")) + @"Arquivos\ExcelExportacao\Modelos\" + strDownloadNome + ".xlsx";

                GetNewFileMemory(ref templateFileName);

                if (File.Exists(fileName))
                    File.Delete(fileName);

                MemoryStream returnedMemoryStream = null;

                using (var file = new FileStream(fileName, FileMode.CreateNew))
                {
                    using (var temp = new FileStream(templateFileName, FileMode.Open))
                    {
                        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                        ExcelPackage package = new ExcelPackage(file, temp);

                        ExcelWorksheet wworksheet = package.Workbook.Worksheets["Data"];

                        wworksheet.View.ShowHeaders = true;
                        wworksheet.View.ShowGridLines = false;
                        //Titulo da planilha. Deixar linha 3 vazias formatadas no modelo excel.
                        wworksheet.Cells[3, 2].Value = tituloPlanilha;

                        int currentRow = 6;
                        int colNameTable = 3;
                        int lastColName = 0;


                        if (grafico != null)
                        {
                            // Preenche cabeçalho:

                            wworksheet.Cells[currentRow, colNameTable].Value = listaTraducaoComponente.FirstOrDefault(a => a.Objeto.Equals("grafico-barras-top-marcas.titulo1")).Texto;
                            colNameTable++;

                            wworksheet.Cells[currentRow, colNameTable].Value = "+/-";
                            colNameTable++;

                            wworksheet.Cells[currentRow, colNameTable].Value = listaTraducaoComponente.FirstOrDefault(a => a.Objeto.Equals("grafico-barras-top-marcas.titulo2")).Texto;
                            colNameTable++;

                            wworksheet.Cells[currentRow, colNameTable].Value = "=";
                            colNameTable++;

                            wworksheet.Cells[currentRow, colNameTable].Value = listaTraducaoComponente.FirstOrDefault(a => a.Objeto.Equals("grafico-barras-top-marcas.titulo3")).Texto;
                            colNameTable++;

                            colNameTable--;
                            lastColName = colNameTable;

                            //Formata cabeçalho
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(3, 6) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Font.Size = 11;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(3, 6) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Font.Name = "Arial Nova Cond";
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(3, 6) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Font.Bold = true;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(3, 6) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(3, 6) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(3, 6) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(3, 6) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(3, 6) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(3, 6) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                            colNameTable = 2;
                            currentRow++;


                            ////Preenche lateral

                            foreach (var item in grafico)
                            {
                                wworksheet.Cells[currentRow, colNameTable].Value = item.DescMarca;
                                Color colour = ColorTranslator.FromHtml(item.CorSiteMarca);
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(colour);
                                currentRow++;
                            }

                            currentRow--;
                            var lastRow = currentRow;
                            colNameTable = 3;
                            currentRow = 6;


                            foreach (var item in grafico)
                            {
                                currentRow++;
                                wworksheet.Cells[currentRow, colNameTable].Value = item.ShareDesejo;
                                if (item.ShareDesejo < -20 || item.ShareDesejo > 20)
                                {
                                    Color colour = ColorTranslator.FromHtml(item.CorSiteMarca);
                                    wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(colour);
                                }

                            }

                            colNameTable++;
                            wworksheet.Column(colNameTable).Width = 5;
                            wworksheet.Cells[currentRow, colNameTable].Value = "";


                            currentRow = 6;
                            colNameTable++;
                            foreach (var item in grafico)
                            {
                                currentRow++;
                                wworksheet.Cells[currentRow, colNameTable].Value = item.Efeitos;
                                if (item.Efeitos < -20 || item.Efeitos > 20)
                                {
                                    Color colour = ColorTranslator.FromHtml(item.CorSiteMarca);
                                    wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(colour);
                                }
                            }

                            colNameTable++;
                            wworksheet.Column(colNameTable).Width = 5;
                            wworksheet.Cells[currentRow, colNameTable].Value = "";


                            currentRow = 6;
                            colNameTable++;
                            foreach (var item in grafico)
                            {
                                currentRow++;
                                wworksheet.Cells[currentRow, colNameTable].Value = item.Equity;
                                if (item.Equity < -20 || item.Equity > 20)
                                {
                                    Color colour = ColorTranslator.FromHtml(item.CorSiteMarca);
                                    wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(colour);
                                }
                            }



                            colNameTable++;
                            currentRow = 7;

                            lastRow++;
                            currentRow = lastRow - 1;
                            colNameTable--;
                            lastColName = colNameTable;

                            //Formata dados
                            //wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 7) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Numberformat.Format = "0.00";
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 7) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Font.Size = 11;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 7) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Font.Name = "Arial Nova Cond";
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 7) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Font.Bold = true;

                            // Preenche lateral
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 11) + ":" + ExcelUtil.GetCelulaByColumnIndex(3, 11)].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 11) + ":" + ExcelUtil.GetCelulaByColumnIndex(3, 11)].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 11) + ":" + ExcelUtil.GetCelulaByColumnIndex(3, 11)].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 11) + ":" + ExcelUtil.GetCelulaByColumnIndex(3, 11)].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 11) + ":" + ExcelUtil.GetCelulaByColumnIndex(3, 11)].Style.Border.Top.Style = ExcelBorderStyle.Thin;

                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 7) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 7) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 7) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 7) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 7) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                        }
                        wworksheet.Select();

                        using (MemoryStream memoryStream = new MemoryStream())
                        {
                            package.SaveAs(memoryStream);

                            memoryStream.Position = 0;

                            returnedMemoryStream = memoryStream;
                        }
                    }
                }

                return File.ReadAllBytes(@"" + fileName);
            }

            catch (Exception)
            {
                throw;
            }
            finally
            {
                File.Delete(fileName);
                File.Delete(templateFileName);
            }
        }


        public byte[] DownloadDashboardSixGraficoBVCEvolutivo(GraficoBVCEvolutivoFull grafico, string tituloPlanilha, List<TraducaoComponente> listaTraducaoComponente, FiltroPadraoExcel filtro)
        {
            string fileName = "";
            string templateFileName = "";
            try
            {
                string strDownloadNome;

                strDownloadNome = "Modelo_Exportação_Gráficos_Padrao";

                string strPath = AdjustPath(HttpContext.Current.Server.MapPath("~")) + @"Arquivos\ExcelExportacao\workspace\";
                fileName = strPath + strDownloadNome + DateTime.Now.ToString("yyyy - MM - dd HH':'mm':'ss':'fff").Replace(" - ", "").Replace(":", "").Replace(" ", "") + ".xlsx";

                templateFileName = AdjustPath(HttpContext.Current.Server.MapPath("~")) + @"Arquivos\ExcelExportacao\Modelos\" + strDownloadNome + ".xlsx";

                GetNewFileMemory(ref templateFileName);

                if (File.Exists(fileName))
                    File.Delete(fileName);

                MemoryStream returnedMemoryStream = null;

                using (var file = new FileStream(fileName, FileMode.CreateNew))
                {
                    using (var temp = new FileStream(templateFileName, FileMode.Open))
                    {
                        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                        ExcelPackage package = new ExcelPackage(file, temp);

                        ExcelWorksheet wworksheet = package.Workbook.Worksheets["Data"];

                        wworksheet.View.ShowHeaders = true;
                        wworksheet.View.ShowGridLines = false;
                        //Titulo da planilha. Deixar linha 3 vazias formatadas no modelo excel.
                        wworksheet.Cells[3, 2].Value = tituloPlanilha;

                        int currentRow = 6;
                        int colNameTable = 3;
                        int lastColName = 0;


                        if (grafico != null)
                        {
                            // Preenche cabeçalho:

                            //wworksheet.Cells[currentRow, colNameTable].Value = listaTraducaoComponente.FirstOrDefault(a => a.Objeto.Equals("top-evolutivo.titulo1")).Texto;
                            //colNameTable++;

                            //wworksheet.Cells[currentRow, colNameTable].Value = listaTraducaoComponente.FirstOrDefault(a => a.Objeto.Equals("top-evolutivo.titulo2")).Texto;
                            //colNameTable++;

                            //wworksheet.Cells[currentRow, colNameTable].Value = listaTraducaoComponente.FirstOrDefault(a => a.Objeto.Equals("top-evolutivo.titulo3")).Texto;
                            //colNameTable++;

                            //wworksheet.Cells[currentRow, colNameTable].Value = listaTraducaoComponente.FirstOrDefault(a => a.Objeto.Equals("top-evolutivo.titulo4")).Texto;
                            //colNameTable++;

                            // Preenche cabeçalho com os periodos:

                            if(filtro.Onda1.FirstOrDefault() == null)
                            {
                                wworksheet.Cells[currentRow, colNameTable].Value = "0";
                            }
                            else
                            wworksheet.Cells[currentRow, colNameTable].Value = filtro.Onda1.FirstOrDefault().DescItem;
                            colNameTable++;

                            if (filtro.Onda2.FirstOrDefault() == null)
                            {
                                wworksheet.Cells[currentRow, colNameTable].Value = "0";
                            }
                            else
                                wworksheet.Cells[currentRow, colNameTable].Value = filtro.Onda2.FirstOrDefault().DescItem;
                            colNameTable++;

                            if (filtro.Onda3.FirstOrDefault() == null)
                            {
                                wworksheet.Cells[currentRow, colNameTable].Value = "0";
                            }
                            else
                                wworksheet.Cells[currentRow, colNameTable].Value = filtro.Onda3.FirstOrDefault().DescItem;
                            colNameTable++;

                            if (filtro.Onda4.FirstOrDefault() == null)
                            {
                                wworksheet.Cells[currentRow, colNameTable].Value = "0";
                            }
                            else
                                wworksheet.Cells[currentRow, colNameTable].Value = filtro.Onda4.FirstOrDefault().DescItem;
                            colNameTable++;

                            colNameTable--;
                            lastColName = colNameTable;

                            //Formata cabeçalho
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 6) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Font.Size = 11;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 6) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Font.Name = "Arial Nova Cond";
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 6) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Font.Bold = true;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 6) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 6) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 6) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 6) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 6) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 6) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                            colNameTable = 2;
                            currentRow++;


                            ////Preenche lateral

                            foreach (var item in grafico.GraficoBVCEvolutivo4)
                            {
                                wworksheet.Cells[currentRow, colNameTable].Value = item.DescMarca;
                                Color colour = ColorTranslator.FromHtml(item.CorSiteMarca);
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(colour);
                                currentRow++;
                            }

                            currentRow--;
                            var lastRow = currentRow;
                            colNameTable = 3;
                            currentRow = 6;


                            foreach (var item in grafico.GraficoBVCEvolutivo1)
                            {
                                currentRow++;
                                wworksheet.Cells[currentRow, colNameTable].Value = item.Perc1;
                                if (item.Perc1 < -20 || item.Perc1 > 20)
                                {
                                    Color colour = ColorTranslator.FromHtml(item.CorSiteMarca);
                                    wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(colour);
                                }

                            }

                            currentRow = 6;
                            colNameTable++;

                            foreach (var item in grafico.GraficoBVCEvolutivo2)
                            {
                                currentRow++;
                                wworksheet.Cells[currentRow, colNameTable].Value = item.Perc1;
                                if (item.Perc1 < -20 || item.Perc1 > 20)
                                {
                                    Color colour = ColorTranslator.FromHtml(item.CorSiteMarca);
                                    wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(colour);
                                }
                            }

                            currentRow = 6;
                            colNameTable++;
                            foreach (var item in grafico.GraficoBVCEvolutivo3)
                            {
                                currentRow++;
                                wworksheet.Cells[currentRow, colNameTable].Value = item.Perc1;
                                if (item.Perc1 < -20 || item.Perc1 > 20)
                                {
                                    Color colour = ColorTranslator.FromHtml(item.CorSiteMarca);
                                    wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(colour);
                                }
                            }

                            currentRow = 6;
                            colNameTable++;
                            foreach (var item in grafico.GraficoBVCEvolutivo4)
                            {
                                currentRow++;
                                wworksheet.Cells[currentRow, colNameTable].Value = item.Perc1;
                                if (item.Perc1 < -20 || item.Perc1 > 20)
                                {
                                    Color colour = ColorTranslator.FromHtml(item.CorSiteMarca);
                                    wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(colour);
                                }
                            }


                            //colNameTable++;
                            //currentRow = 6;

                            //lastRow++;
                            //currentRow = lastRow - 1;
                            //colNameTable--;
                            //lastColName = colNameTable;





                            //Formata dados
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 6) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Font.Size = 11;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 6) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Font.Name = "Arial Nova Cond";
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 6) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Font.Bold = true;

                            // Preenche lateral
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 11) + ":" + ExcelUtil.GetCelulaByColumnIndex(3, 11)].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 11) + ":" + ExcelUtil.GetCelulaByColumnIndex(3, 11)].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 11) + ":" + ExcelUtil.GetCelulaByColumnIndex(3, 11)].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 11) + ":" + ExcelUtil.GetCelulaByColumnIndex(3, 11)].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 11) + ":" + ExcelUtil.GetCelulaByColumnIndex(3, 11)].Style.Border.Top.Style = ExcelBorderStyle.Thin;

                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 6) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 6) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 6) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 6) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 6) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                        }
                        wworksheet.Select();

                        using (MemoryStream memoryStream = new MemoryStream())
                        {
                            package.SaveAs(memoryStream);

                            memoryStream.Position = 0;

                            returnedMemoryStream = memoryStream;
                        }
                    }
                }

                return File.ReadAllBytes(@"" + fileName);
            }

            catch (Exception)
            {
                throw;
            }
            finally
            {
                File.Delete(fileName);
                File.Delete(templateFileName);
            }
        }


        #region DownloadDashboardSeven

        public byte[] DownloadDashboardGraficoComunicacaoRecall(List<GraficoComunicacaoRecall> grafico, string tituloPlanilha, List<TraducaoComponente> listaTraducaoComponente)
        {
            string fileName = "";
            string templateFileName = "";
            try
            {
                string strDownloadNome;

                strDownloadNome = "Modelo_Exportação_Gráficos_Padrao";

                string strPath = AdjustPath(HttpContext.Current.Server.MapPath("~")) + @"Arquivos\ExcelExportacao\workspace\";
                fileName = strPath + strDownloadNome + DateTime.Now.ToString("yyyy - MM - dd HH':'mm':'ss':'fff").Replace(" - ", "").Replace(":", "").Replace(" ", "") + ".xlsx";

                templateFileName = AdjustPath(HttpContext.Current.Server.MapPath("~")) + @"Arquivos\ExcelExportacao\Modelos\" + strDownloadNome + ".xlsx";

                GetNewFileMemory(ref templateFileName);

                if (File.Exists(fileName))
                    File.Delete(fileName);

                MemoryStream returnedMemoryStream = null;

                using (var file = new FileStream(fileName, FileMode.CreateNew))
                {
                    using (var temp = new FileStream(templateFileName, FileMode.Open))
                    {
                        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                        ExcelPackage package = new ExcelPackage(file, temp);

                        ExcelWorksheet wworksheet = package.Workbook.Worksheets["Data"];

                        wworksheet.View.ShowHeaders = true;
                        wworksheet.View.ShowGridLines = false;
                        //Titulo da planilha. Deixar linha 3 vazias formatadas no modelo excel.
                        wworksheet.Cells[3, 2].Value = tituloPlanilha;

                        int currentRow = 6;
                        int colNameTable = 3;
                        int lastColName = 0;


                        if (grafico != null)
                        {
                            // Preenche cabeçalho:
                            Color colourGray = ColorTranslator.FromHtml("#BFBFBF");

                            wworksheet.Cells[currentRow, colNameTable].Value = listaTraducaoComponente.FirstOrDefault(a => a.Objeto.Equals("recall-coluna1.titulo")).Texto;
                            wworksheet.Cells[currentRow + 1, colNameTable].Value = listaTraducaoComponente.FirstOrDefault(a => a.Objeto.Equals("recall-coluna1.SubTitulo")).Texto;
                            wworksheet.Cells[currentRow + 1, colNameTable].Style.Font.Color.SetColor(colourGray);
                            colNameTable++;

                            wworksheet.Cells[currentRow, colNameTable].Value = listaTraducaoComponente.FirstOrDefault(a => a.Objeto.Equals("recall-coluna2.titulo")).Texto;
                            wworksheet.Cells[currentRow + 1, colNameTable].Value = listaTraducaoComponente.FirstOrDefault(a => a.Objeto.Equals("recall-coluna2.SubTitulo")).Texto;
                            wworksheet.Cells[currentRow + 1, colNameTable].Style.Font.Color.SetColor(colourGray);
                            colNameTable++;

                            wworksheet.Cells[currentRow, colNameTable].Value = listaTraducaoComponente.FirstOrDefault(a => a.Objeto.Equals("recall-coluna3.titulo")).Texto;
                            wworksheet.Cells[currentRow + 1, colNameTable].Value = listaTraducaoComponente.FirstOrDefault(a => a.Objeto.Equals("recall-coluna3.SubTitulo")).Texto;
                            wworksheet.Cells[currentRow + 1, colNameTable].Style.Font.Color.SetColor(colourGray);
                            colNameTable++;

                            colNameTable--;
                            lastColName = colNameTable;

                            //Formata cabeçalho
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 6) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Font.Size = 11;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 6) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Font.Name = "Arial Nova Cond";
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 6) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Font.Bold = true;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 6) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 6) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 6) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 6) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 6) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 6) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                            colNameTable = 2;
                            currentRow++;
                            currentRow++;

                            ////Preenche lateral                        
                            wworksheet.Cells[currentRow, colNameTable].Value = listaTraducaoComponente.FirstOrDefault(a => a.Objeto.Equals("dashboard-comunicacao-titulo1")).Texto;
                            Color colour = ColorTranslator.FromHtml("#4056B7");
                            wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(colour);
                            currentRow++;

                            wworksheet.Cells[currentRow, colNameTable].Value = listaTraducaoComponente.FirstOrDefault(a => a.Objeto.Equals("dashboard-comunicacao-titulo2")).Texto;
                            Color colour2 = ColorTranslator.FromHtml("#595959");
                            wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(colour2);
                            currentRow++;

                            wworksheet.Cells[currentRow, colNameTable].Value = "Base";
                            currentRow++;

                            currentRow--;
                            var lastRow = currentRow;
                            colNameTable = 3;
                            currentRow = 7;
                            currentRow++;

                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.FirstOrDefault().VisibilidadeMarca;
                            if (grafico.FirstOrDefault().VisibilidadeTesteSIGMarca == "MENOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Red);
                            if (grafico.FirstOrDefault().VisibilidadeTesteSIGMarca == "MAIOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Blue);
                            currentRow++;

                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.FirstOrDefault().VisibilidadeNorma;
                            currentRow++;

                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.FirstOrDefault().VisibilidadeBase;
                            if (!string.IsNullOrWhiteSpace(grafico.FirstOrDefault().VisibilidadeBaseMinima))
                            {
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Red);
                            }


                            colNameTable++;
                            currentRow = 7;
                            currentRow++;

                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.FirstOrDefault().LinkageMarca;
                            if (grafico.FirstOrDefault().LinkageTesteSIGMarca == "MENOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Red);
                            if (grafico.FirstOrDefault().LinkageTesteSIGMarca == "MAIOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Blue);
                            currentRow++;

                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.FirstOrDefault().LinkageNorma;
                            currentRow++;

                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.FirstOrDefault().LinkageBase;
                            if (!string.IsNullOrWhiteSpace(grafico.FirstOrDefault().LinkageBaseMinima))
                            {
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Red);
                            }

                            colNameTable++;
                            currentRow = 7;
                            currentRow++;

                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.FirstOrDefault().RecalMarca;
                            if (grafico.FirstOrDefault().RecalTesteSIGMarca == "MENOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Red);
                            if (grafico.FirstOrDefault().RecalTesteSIGMarca == "MAIOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Blue);
                            currentRow++;

                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.FirstOrDefault().RecalNorma;
                            currentRow++;

                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.FirstOrDefault().RecalBase;
                            if (!string.IsNullOrWhiteSpace(grafico.FirstOrDefault().RecalBaseMinima))
                            {
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Red);
                            }



                            colNameTable++;
                            currentRow = 5;

                            lastRow++;
                            currentRow = lastRow;
                            colNameTable--;
                            lastColName = colNameTable;


                            //Formata dados
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 6) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Font.Size = 11;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 6) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Font.Name = "Arial Nova Cond";
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 6) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Font.Bold = true;

                            // Preenche lateral
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 11) + ":" + ExcelUtil.GetCelulaByColumnIndex(3, 11)].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 11) + ":" + ExcelUtil.GetCelulaByColumnIndex(3, 11)].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 11) + ":" + ExcelUtil.GetCelulaByColumnIndex(3, 11)].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 11) + ":" + ExcelUtil.GetCelulaByColumnIndex(3, 11)].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 11) + ":" + ExcelUtil.GetCelulaByColumnIndex(3, 11)].Style.Border.Top.Style = ExcelBorderStyle.Thin;

                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 6) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 6) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 6) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 6) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 6) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                        }
                        wworksheet.Select();

                        using (MemoryStream memoryStream = new MemoryStream())
                        {
                            package.SaveAs(memoryStream);

                            memoryStream.Position = 0;

                            returnedMemoryStream = memoryStream;
                        }
                    }
                }

                return File.ReadAllBytes(@"" + fileName);
            }

            catch (Exception)
            {
                throw;
            }
            finally
            {
                File.Delete(fileName);
                File.Delete(templateFileName);
            }
        }


        public byte[] DownloadDashboardGraficoComunicacaoVisto(List<GraficoComunicacaoVisto> grafico, string tituloPlanilha, List<TraducaoComponente> listaTraducaoComponente)
        {
            string fileName = "";
            string templateFileName = "";
            try
            {
                string strDownloadNome;

                strDownloadNome = "Modelo_Exportação_Gráficos_Padrao";

                string strPath = AdjustPath(HttpContext.Current.Server.MapPath("~")) + @"Arquivos\ExcelExportacao\workspace\";
                fileName = strPath + strDownloadNome + DateTime.Now.ToString("yyyy - MM - dd HH':'mm':'ss':'fff").Replace(" - ", "").Replace(":", "").Replace(" ", "") + ".xlsx";

                templateFileName = AdjustPath(HttpContext.Current.Server.MapPath("~")) + @"Arquivos\ExcelExportacao\Modelos\" + strDownloadNome + ".xlsx";

                GetNewFileMemory(ref templateFileName);

                if (File.Exists(fileName))
                    File.Delete(fileName);

                MemoryStream returnedMemoryStream = null;

                using (var file = new FileStream(fileName, FileMode.CreateNew))
                {
                    using (var temp = new FileStream(templateFileName, FileMode.Open))
                    {
                        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                        ExcelPackage package = new ExcelPackage(file, temp);

                        ExcelWorksheet wworksheet = package.Workbook.Worksheets["Data"];

                        wworksheet.View.ShowHeaders = true;
                        wworksheet.View.ShowGridLines = false;
                        //Titulo da planilha. Deixar linha 3 vazias formatadas no modelo excel.
                        wworksheet.Cells[3, 2].Value = tituloPlanilha;

                        int currentRow = 6;
                        int colNameTable = 3;
                        int lastColName = 0;


                        if (grafico != null)
                        {
                            // Preenche cabeçalho:


                            wworksheet.Cells[currentRow, colNameTable].Value = "Perc";
                            colNameTable++;
                            wworksheet.Cells[currentRow, colNameTable].Value = "Base";
                            colNameTable++;


                            colNameTable--;
                            lastColName = colNameTable;

                            //Formata cabeçalho
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 6) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Font.Size = 11;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 6) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Font.Name = "Arial Nova Cond";
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 6) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Font.Bold = true;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 6) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 6) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 6) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 6) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 6) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 6) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                            colNameTable = 2;
                            currentRow++;


                            ////Preenche lateral

                            foreach (var item in grafico)
                            {
                                wworksheet.Cells[currentRow, colNameTable].Value = item.Descricao;
                                currentRow++;
                            }

                            currentRow--;
                            var lastRow = currentRow;
                            colNameTable = 3;
                            currentRow = 6;


                            foreach (var item in grafico)
                            {
                                currentRow++;
                                wworksheet.Cells[currentRow, colNameTable].Value = item.perc;
                                //Color colour = ColorTranslator.FromHtml("blue");
                                //wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(colour);

                                wworksheet.Cells[currentRow, colNameTable + 1].Value = item.QtdeAbs;
                                if (!string.IsNullOrWhiteSpace(item.BaseMinima))
                                {
                                    wworksheet.Cells[currentRow, colNameTable + 1].Style.Font.Color.SetColor(Color.Red);
                                }
                                //Color colour2 = ColorTranslator.FromHtml("#BFBFBF");
                                //wworksheet.Cells[currentRow, colNameTable + 1].Style.Font.Color.SetColor(colour2);

                            }


                            //colNameTable++;
                            //currentRow = 6;

                            //lastRow++;
                            //currentRow = lastRow - 1;
                            //colNameTable--;
                            //lastColName = colNameTable;



                            //Formata dados
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 6) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Font.Size = 11;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 6) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Font.Name = "Arial Nova Cond";
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 6) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Font.Bold = true;

                            // Preenche lateral
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 11) + ":" + ExcelUtil.GetCelulaByColumnIndex(3, 11)].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 11) + ":" + ExcelUtil.GetCelulaByColumnIndex(3, 11)].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 11) + ":" + ExcelUtil.GetCelulaByColumnIndex(3, 11)].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 11) + ":" + ExcelUtil.GetCelulaByColumnIndex(3, 11)].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 11) + ":" + ExcelUtil.GetCelulaByColumnIndex(3, 11)].Style.Border.Top.Style = ExcelBorderStyle.Thin;

                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 6) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 6) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 6) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 6) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 6) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                        }
                        wworksheet.Select();

                        using (MemoryStream memoryStream = new MemoryStream())
                        {
                            package.SaveAs(memoryStream);

                            memoryStream.Position = 0;

                            returnedMemoryStream = memoryStream;
                        }
                    }
                }

                return File.ReadAllBytes(@"" + fileName);
            }

            catch (Exception)
            {
                throw;
            }
            finally
            {
                File.Delete(fileName);
                File.Delete(templateFileName);
            }
        }


        public byte[] DownloadDashboardGraficoComunicacaoSource(List<GraficoComunicacaoVisto> grafico, string tituloPlanilha, List<TraducaoComponente> listaTraducaoComponente)
        {
            string fileName = "";
            string templateFileName = "";
            try
            {
                string strDownloadNome;

                strDownloadNome = "Modelo_Exportação_Gráficos_Padrao";

                string strPath = AdjustPath(HttpContext.Current.Server.MapPath("~")) + @"Arquivos\ExcelExportacao\workspace\";
                fileName = strPath + strDownloadNome + DateTime.Now.ToString("yyyy - MM - dd HH':'mm':'ss':'fff").Replace(" - ", "").Replace(":", "").Replace(" ", "") + ".xlsx";

                templateFileName = AdjustPath(HttpContext.Current.Server.MapPath("~")) + @"Arquivos\ExcelExportacao\Modelos\" + strDownloadNome + ".xlsx";

                GetNewFileMemory(ref templateFileName);

                if (File.Exists(fileName))
                    File.Delete(fileName);

                MemoryStream returnedMemoryStream = null;

                using (var file = new FileStream(fileName, FileMode.CreateNew))
                {
                    using (var temp = new FileStream(templateFileName, FileMode.Open))
                    {
                        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                        ExcelPackage package = new ExcelPackage(file, temp);

                        ExcelWorksheet wworksheet = package.Workbook.Worksheets["Data"];

                        wworksheet.View.ShowHeaders = true;
                        wworksheet.View.ShowGridLines = false;
                        //Titulo da planilha. Deixar linha 3 vazias formatadas no modelo excel.
                        wworksheet.Cells[3, 2].Value = tituloPlanilha;

                        int currentRow = 6;
                        int colNameTable = 3;
                        int lastColName = 0;


                        if (grafico != null)
                        {
                            // Preenche cabeçalho:


                            wworksheet.Cells[currentRow, colNameTable].Value = "Perc";
                            colNameTable++;
                            wworksheet.Cells[currentRow, colNameTable].Value = "Base";
                            colNameTable++;


                            colNameTable--;
                            lastColName = colNameTable;

                            //Formata cabeçalho
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 6) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Font.Size = 11;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 6) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Font.Name = "Arial Nova Cond";
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 6) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Font.Bold = true;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 6) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 6) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 6) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 6) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 6) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 6) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                            colNameTable = 2;
                            currentRow++;


                            ////Preenche lateral

                            foreach (var item in grafico)
                            {
                                wworksheet.Cells[currentRow, colNameTable].Value = item.Descricao;
                                currentRow++;
                            }

                            currentRow--;
                            var lastRow = currentRow;
                            colNameTable = 3;
                            currentRow = 6;


                            foreach (var item in grafico)
                            {
                                currentRow++;
                                wworksheet.Cells[currentRow, colNameTable].Value = item.perc;
                                //Color colour = ColorTranslator.FromHtml("blue");
                                //wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(colour);

                                wworksheet.Cells[currentRow, colNameTable + 1].Value = item.QtdeAbs;
                                if (!string.IsNullOrWhiteSpace(item.BaseMinima))
                                {
                                    wworksheet.Cells[currentRow, colNameTable + 1].Style.Font.Color.SetColor(Color.Red);
                                }
                                //Color colour2 = ColorTranslator.FromHtml("#BFBFBF");
                                //wworksheet.Cells[currentRow, colNameTable + 1].Style.Font.Color.SetColor(colour2);

                            }


                            //colNameTable++;
                            //currentRow = 6;

                            //lastRow++;
                            //currentRow = lastRow - 1;
                            //colNameTable--;
                            //lastColName = colNameTable;



                            //Formata dados
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 6) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Font.Size = 11;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 6) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Font.Name = "Arial Nova Cond";
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 6) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Font.Bold = true;

                            // Preenche lateral
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 11) + ":" + ExcelUtil.GetCelulaByColumnIndex(3, 11)].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 11) + ":" + ExcelUtil.GetCelulaByColumnIndex(3, 11)].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 11) + ":" + ExcelUtil.GetCelulaByColumnIndex(3, 11)].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 11) + ":" + ExcelUtil.GetCelulaByColumnIndex(3, 11)].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 11) + ":" + ExcelUtil.GetCelulaByColumnIndex(3, 11)].Style.Border.Top.Style = ExcelBorderStyle.Thin;

                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 6) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 6) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 6) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 6) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 6) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                        }
                        wworksheet.Select();

                        using (MemoryStream memoryStream = new MemoryStream())
                        {
                            package.SaveAs(memoryStream);

                            memoryStream.Position = 0;

                            returnedMemoryStream = memoryStream;
                        }
                    }
                }

                return File.ReadAllBytes(@"" + fileName);
            }

            catch (Exception)
            {
                throw;
            }
            finally
            {
                File.Delete(fileName);
                File.Delete(templateFileName);
            }
        }


        public byte[] DownloadDashboardGraficoComunicacaoDiagnostico(List<GraficoComunicacaoDiagnostico> grafico, string tituloPlanilha, List<TraducaoComponente> listaTraducaoComponente)
        {
            string fileName = "";
            string templateFileName = "";
            try
            {
                string strDownloadNome;

                strDownloadNome = "Modelo_Exportação_Gráficos_Padrao";

                string strPath = AdjustPath(HttpContext.Current.Server.MapPath("~")) + @"Arquivos\ExcelExportacao\workspace\";
                fileName = strPath + strDownloadNome + DateTime.Now.ToString("yyyy - MM - dd HH':'mm':'ss':'fff").Replace(" - ", "").Replace(":", "").Replace(" ", "") + ".xlsx";

                templateFileName = AdjustPath(HttpContext.Current.Server.MapPath("~")) + @"Arquivos\ExcelExportacao\Modelos\" + strDownloadNome + ".xlsx";

                GetNewFileMemory(ref templateFileName);

                if (File.Exists(fileName))
                    File.Delete(fileName);

                MemoryStream returnedMemoryStream = null;

                using (var file = new FileStream(fileName, FileMode.CreateNew))
                {
                    using (var temp = new FileStream(templateFileName, FileMode.Open))
                    {
                        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                        ExcelPackage package = new ExcelPackage(file, temp);

                        ExcelWorksheet wworksheet = package.Workbook.Worksheets["Data"];

                        wworksheet.View.ShowHeaders = true;
                        wworksheet.View.ShowGridLines = false;
                        //Titulo da planilha. Deixar linha 3 vazias formatadas no modelo excel.
                        wworksheet.Cells[3, 2].Value = tituloPlanilha;

                        int currentRow = 6;
                        int colNameTable = 3;
                        int lastColName = 0;


                        if (grafico != null)
                        {
                            // Preenche cabeçalho:

                            Color colourT1 = ColorTranslator.FromHtml("#4056B7");
                            wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(colourT1);
                            wworksheet.Cells[currentRow, colNameTable].Value = listaTraducaoComponente.FirstOrDefault(a => a.Objeto.Equals("dashboard-comunicacao-titulo1")).Texto;
                            colNameTable++;

                            Color colouT2 = ColorTranslator.FromHtml("#BDBDBD");
                            wworksheet.Cells[currentRow, colNameTable + 1].Style.Font.Color.SetColor(colouT2);
                            wworksheet.Cells[currentRow, colNameTable].Value = listaTraducaoComponente.FirstOrDefault(a => a.Objeto.Equals("dashboard-comunicacao-titulo2")).Texto; ;
                            colNameTable++;
                            wworksheet.Cells[currentRow, colNameTable].Value = listaTraducaoComponente.FirstOrDefault(a => a.Objeto.Equals("dashboard-diagnostico-propaganda.GAP")).Texto;
                            colNameTable++;

                            Color colouT3 = ColorTranslator.FromHtml("#BDBDBD");
                            wworksheet.Cells[currentRow, colNameTable + 2].Style.Font.Color.SetColor(colouT3);
                            wworksheet.Cells[currentRow, colNameTable].Value = "Base";
                            colNameTable++;



                            colNameTable--;
                            lastColName = colNameTable;

                            //Formata cabeçalho
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 6) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Font.Size = 11;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 6) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Font.Name = "Arial Nova Cond";
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 6) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Font.Bold = true;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 6) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 6) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 6) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 6) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 6) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 6) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                            colNameTable = 2;
                            currentRow++;


                            ////Preenche lateral

                            foreach (var item in grafico)
                            {
                                wworksheet.Cells[currentRow, colNameTable].Value = item.Descricao;
                                currentRow++;
                            }

                            currentRow--;
                            var lastRow = currentRow;
                            colNameTable = 3;
                            currentRow = 6;


                            foreach (var item in grafico)
                            {
                                currentRow++;
                                wworksheet.Cells[currentRow, colNameTable].Value = item.Perc;
                                //Color colour = ColorTranslator.FromHtml("#4056B7");
                                //wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(colour);

                                if (item.TesteSig == "MENOR")
                                    wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Red);
                                if (item.TesteSig == "MAIOR")
                                    wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Blue);

                                if (!string.IsNullOrWhiteSpace(item.BaseMinima))
                                {
                                    wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Red);
                                }

                                wworksheet.Cells[currentRow, colNameTable + 1].Value = item.PercNorma;
                                Color colour2 = ColorTranslator.FromHtml("#BDBDBD");
                                wworksheet.Cells[currentRow, colNameTable + 1].Style.Font.Color.SetColor(colour2);

                                wworksheet.Cells[currentRow, colNameTable + 2].Value = item.GAP;


                                if (!string.IsNullOrWhiteSpace(item.BaseMinima))
                                {
                                    wworksheet.Cells[currentRow, colNameTable + 3].Style.Font.Color.SetColor(Color.Red);
                                }
                                wworksheet.Cells[currentRow, colNameTable + 3].Value = item.Base;

                            }

                            //colNameTable++;
                            //currentRow = 6;

                            //lastRow++;
                            //currentRow = lastRow - 1;
                            //colNameTable--;
                            //lastColName = colNameTable;



                            //Formata dados
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 6) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Font.Size = 11;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 6) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Font.Name = "Arial Nova Cond";
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 6) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Font.Bold = true;

                            // Preenche lateral
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 11) + ":" + ExcelUtil.GetCelulaByColumnIndex(3, 11)].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 11) + ":" + ExcelUtil.GetCelulaByColumnIndex(3, 11)].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 11) + ":" + ExcelUtil.GetCelulaByColumnIndex(3, 11)].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 11) + ":" + ExcelUtil.GetCelulaByColumnIndex(3, 11)].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 11) + ":" + ExcelUtil.GetCelulaByColumnIndex(3, 11)].Style.Border.Top.Style = ExcelBorderStyle.Thin;

                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 6) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 6) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 6) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 6) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 6) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                        }
                        wworksheet.Select();

                        using (MemoryStream memoryStream = new MemoryStream())
                        {
                            package.SaveAs(memoryStream);

                            memoryStream.Position = 0;

                            returnedMemoryStream = memoryStream;
                        }
                    }
                }

                return File.ReadAllBytes(@"" + fileName);
            }

            catch (Exception)
            {
                throw;
            }
            finally
            {
                File.Delete(fileName);
                File.Delete(templateFileName);
            }
        }

        #endregion


        #region DownloadDashboardTwo
        public byte[] DownloadDashboardEightComparativoMarcas(GraficoColunasFullLoad grafico, string tituloPlanilha, List<TraducaoComponente> listaTraducaoComponente)
        {
            string fileName = "";
            string templateFileName = "";
            try
            {
                string strDownloadNome;

                strDownloadNome = "Modelo_Exportação_Gráficos_Padrao";

                string strPath = AdjustPath(HttpContext.Current.Server.MapPath("~")) + @"Arquivos\ExcelExportacao\workspace\";
                fileName = strPath + strDownloadNome + DateTime.Now.ToString("yyyy - MM - dd HH':'mm':'ss':'fff").Replace(" - ", "").Replace(":", "").Replace(" ", "") + ".xlsx";

                templateFileName = AdjustPath(HttpContext.Current.Server.MapPath("~")) + @"Arquivos\ExcelExportacao\Modelos\" + strDownloadNome + ".xlsx";

                GetNewFileMemory(ref templateFileName);

                if (File.Exists(fileName))
                    File.Delete(fileName);

                MemoryStream returnedMemoryStream = null;

                using (var file = new FileStream(fileName, FileMode.CreateNew))
                {
                    using (var temp = new FileStream(templateFileName, FileMode.Open))
                    {
                        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                        ExcelPackage package = new ExcelPackage(file, temp);

                        ExcelWorksheet wworksheet = package.Workbook.Worksheets["Data"];

                        wworksheet.View.ShowHeaders = true;
                        wworksheet.View.ShowGridLines = false;
                        //Titulo da planilha. Deixar linha 3 vazias formatadas no modelo excel.
                        wworksheet.Cells[3, 2].Value = tituloPlanilha;

                        int currentRow = 6;
                        int colNameTable = 3;
                        int lastColName = 0;


                        if (grafico != null)
                        {
                            // Preenche cabeçalho com os periodos:
                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoColunas1.DescMarca;
                            colNameTable++;
                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoColunas2.DescMarca;
                            colNameTable++;

                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoColunas3.DescMarca;
                            colNameTable++;

                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoColunas4.DescMarca;
                            colNameTable++;

                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoColunas5.DescMarca;
                            colNameTable++;

                            colNameTable--;
                            lastColName = colNameTable;

                            //Formata cabeçalho
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(3, 6) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Font.Size = 11;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(3, 6) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Font.Name = "Arial Nova Cond";
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(3, 6) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Font.Bold = true;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(3, 6) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(3, 6) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(3, 6) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(3, 6) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(3, 6) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(3, 6) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                            colNameTable = 2;
                            currentRow++;


                            ////Preenche lateral
                            //wworksheet.Cells[currentRow, colNameTable].Value = "Total Awareness";
                            //currentRow++;
                            //wworksheet.Cells[currentRow, colNameTable].Value = "Prompeted";
                            //currentRow++;
                            //wworksheet.Cells[currentRow, colNameTable].Value = "Spontaneous - OM";
                            //currentRow++;
                            //wworksheet.Cells[currentRow, colNameTable].Value = "Awareness - TOM";
                            //currentRow++;
                            //wworksheet.Cells[currentRow, colNameTable].Value = "Base";

                            wworksheet.Cells[currentRow, colNameTable].Value = listaTraducaoComponente.FirstOrDefault(a => a.Objeto.Equals("dashboard-eight-leg-Awareness")).Texto;
                            currentRow++;

                            wworksheet.Cells[currentRow, colNameTable].Value = listaTraducaoComponente.FirstOrDefault(a => a.Objeto.Equals("dashboard-eight-leg-Spontaneous")).Texto;
                            currentRow++;

                            wworksheet.Cells[currentRow, colNameTable].Value = listaTraducaoComponente.FirstOrDefault(a => a.Objeto.Equals("dashboard-eight-leg-Prompeted")).Texto;
                            currentRow++;

                            wworksheet.Cells[currentRow, colNameTable].Value = listaTraducaoComponente.FirstOrDefault(a => a.Objeto.Equals("dashboard-eight-leg-total-Awareness")).Texto;
                            currentRow++;

                            wworksheet.Cells[currentRow, colNameTable].Value = listaTraducaoComponente.FirstOrDefault(a => a.Objeto.Equals("dashboard-eight-leg-outros")).Texto;
                            currentRow++;
                            wworksheet.Cells[currentRow, colNameTable].Value = "Base";

                            currentRow--;
                            var lastRow = currentRow;
                            colNameTable = 3;
                            currentRow = 7;



                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoColunas1.PercTotal;
                            if (grafico.GraficoColunas1.TesteSigTotal == "MENOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Red);
                            if (grafico.GraficoColunas1.TesteSigTotal == "MAIOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Blue);

                            currentRow++;
                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoColunas1.PercPrompeted;
                            if (grafico.GraficoColunas1.TesteSIGPrompeted == "MENOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Red);
                            if (grafico.GraficoColunas1.TesteSIGPrompeted == "MAIOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Blue);

                            currentRow++;
                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoColunas1.PercOM;
                            if (grafico.GraficoColunas1.TesteSIGOM == "MENOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Red);
                            if (grafico.GraficoColunas1.TesteSIGOM == "MAIOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Blue);

                            currentRow++;
                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoColunas1.PercTOM;
                            if (grafico.GraficoColunas1.TesteSIGTOM == "MENOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Red);
                            if (grafico.GraficoColunas1.TesteSIGTOM == "MAIOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Blue);
                            currentRow++;

                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoColunas1.PercOutros;
                            if (grafico.GraficoColunas1.TesteSigOutros == "MENOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Red);
                            if (grafico.GraficoColunas1.TesteSigOutros == "MAIOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Blue);
                            currentRow++;

                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoColunas1.BaseAbs;
                            colNameTable++;
                            currentRow = 7;

                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoColunas2.PercTotal;
                            if (grafico.GraficoColunas2.TesteSigTotal == "MENOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Red);
                            if (grafico.GraficoColunas2.TesteSigTotal == "MAIOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Blue);
                            currentRow++;

                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoColunas2.PercPrompeted;
                            if (grafico.GraficoColunas2.TesteSIGPrompeted == "MENOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Red);
                            if (grafico.GraficoColunas2.TesteSIGPrompeted == "MAIOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Blue);
                            currentRow++;

                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoColunas2.PercOM;
                            if (grafico.GraficoColunas2.TesteSIGOM == "MENOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Red);
                            if (grafico.GraficoColunas2.TesteSIGOM == "MAIOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Blue);
                            currentRow++;

                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoColunas2.PercTOM;
                            if (grafico.GraficoColunas2.TesteSIGTOM == "MENOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Red);
                            if (grafico.GraficoColunas2.TesteSIGTOM == "MAIOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Blue);
                            currentRow++;

                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoColunas2.PercOutros;
                            if (grafico.GraficoColunas2.TesteSigOutros == "MENOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Red);
                            if (grafico.GraficoColunas2.TesteSigOutros == "MAIOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Blue);
                            currentRow++;

                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoColunas2.BaseAbs;
                            colNameTable++;
                            currentRow = 7;

                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoColunas3.PercTotal;
                            if (grafico.GraficoColunas3.TesteSigTotal == "MENOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Red);
                            if (grafico.GraficoColunas3.TesteSigTotal == "MAIOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Blue);
                            currentRow++;

                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoColunas3.PercPrompeted;
                            if (grafico.GraficoColunas3.TesteSIGPrompeted == "MENOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Red);
                            if (grafico.GraficoColunas3.TesteSIGPrompeted == "MAIOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Blue);
                            currentRow++;

                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoColunas3.PercOM;
                            if (grafico.GraficoColunas3.TesteSIGOM == "MENOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Red);
                            if (grafico.GraficoColunas3.TesteSIGOM == "MAIOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Blue);
                            currentRow++;

                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoColunas3.PercTOM;
                            if (grafico.GraficoColunas3.TesteSIGTOM == "MENOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Red);
                            if (grafico.GraficoColunas3.TesteSIGTOM == "MAIOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Blue);
                            currentRow++;

                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoColunas3.PercOutros;
                            if (grafico.GraficoColunas3.TesteSigOutros == "MENOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Red);
                            if (grafico.GraficoColunas3.TesteSigOutros == "MAIOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Blue);
                            currentRow++;

                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoColunas3.BaseAbs;
                            colNameTable++;
                            currentRow = 7;

                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoColunas4.PercTotal;
                            if (grafico.GraficoColunas4.TesteSigTotal == "MENOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Red);
                            if (grafico.GraficoColunas4.TesteSigTotal == "MAIOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Blue);
                            currentRow++;

                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoColunas4.PercPrompeted;
                            if (grafico.GraficoColunas4.TesteSIGPrompeted == "MENOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Red);
                            if (grafico.GraficoColunas4.TesteSIGPrompeted == "MAIOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Blue);
                            currentRow++;

                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoColunas4.PercOM;
                            if (grafico.GraficoColunas4.TesteSIGOM == "MENOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Red);
                            if (grafico.GraficoColunas4.TesteSIGOM == "MAIOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Blue);
                            currentRow++;

                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoColunas4.PercTOM;
                            if (grafico.GraficoColunas4.TesteSIGTOM == "MENOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Red);
                            if (grafico.GraficoColunas4.TesteSIGTOM == "MAIOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Blue);
                            currentRow++;

                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoColunas4.PercOutros;
                            if (grafico.GraficoColunas4.TesteSigOutros == "MENOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Red);
                            if (grafico.GraficoColunas4.TesteSigOutros == "MAIOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Blue);
                            currentRow++;

                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoColunas4.BaseAbs;
                            colNameTable++;
                            currentRow = 7;

                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoColunas5.PercTotal;
                            if (grafico.GraficoColunas5.TesteSigTotal == "MENOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Red);
                            if (grafico.GraficoColunas5.TesteSigTotal == "MAIOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Blue);
                            currentRow++;

                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoColunas5.PercPrompeted;
                            if (grafico.GraficoColunas5.TesteSIGPrompeted == "MENOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Red);
                            if (grafico.GraficoColunas5.TesteSIGPrompeted == "MAIOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Blue);
                            currentRow++;

                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoColunas5.PercOM;
                            if (grafico.GraficoColunas5.TesteSIGOM == "MENOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Red);
                            if (grafico.GraficoColunas5.TesteSIGOM == "MAIOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Blue);
                            currentRow++;

                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoColunas5.PercTOM;
                            if (grafico.GraficoColunas5.TesteSIGTOM == "MENOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Red);
                            if (grafico.GraficoColunas5.TesteSIGTOM == "MAIOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Blue);
                            currentRow++;

                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoColunas5.PercOutros;
                            if (grafico.GraficoColunas5.TesteSigOutros == "MENOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Red);
                            if (grafico.GraficoColunas5.TesteSigOutros == "MAIOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Blue);
                            currentRow++;

                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoColunas5.BaseAbs;
                            colNameTable++;
                            currentRow = 7;


                            lastRow++;
                            currentRow = lastRow;
                            colNameTable--;
                            lastColName = colNameTable;

                            //Formata dados
                            //wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 7) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Numberformat.Format = "0.00";
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 7) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Font.Size = 11;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 7) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Font.Name = "Arial Nova Cond";
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 7) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Font.Bold = true;

                            // Preenche lateral
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 11) + ":" + ExcelUtil.GetCelulaByColumnIndex(3, 11)].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 11) + ":" + ExcelUtil.GetCelulaByColumnIndex(3, 11)].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 11) + ":" + ExcelUtil.GetCelulaByColumnIndex(3, 11)].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 11) + ":" + ExcelUtil.GetCelulaByColumnIndex(3, 11)].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 11) + ":" + ExcelUtil.GetCelulaByColumnIndex(3, 11)].Style.Border.Top.Style = ExcelBorderStyle.Thin;

                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 7) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 7) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 7) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 7) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 7) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                        }
                        wworksheet.Select();

                        using (MemoryStream memoryStream = new MemoryStream())
                        {
                            package.SaveAs(memoryStream);

                            memoryStream.Position = 0;

                            returnedMemoryStream = memoryStream;
                        }
                    }
                }

                return File.ReadAllBytes(@"" + fileName);
            }

            catch (Exception)
            {
                throw;
            }
            finally
            {
                File.Delete(fileName);
                File.Delete(templateFileName);
            }
        }

        public byte[] DownloadDashboardEightEvolutivoMarcas(GraficoColunasFullLoad grafico, string tituloPlanilha, List<TraducaoComponente> listaTraducaoComponente, FiltroPadraoExcel filtro)
        {
            string fileName = "";
            string templateFileName = "";
            try
            {
                string strDownloadNome;

                strDownloadNome = "Modelo_Exportação_Gráficos_Padrao";

                string strPath = AdjustPath(HttpContext.Current.Server.MapPath("~")) + @"Arquivos\ExcelExportacao\workspace\";
                fileName = strPath + strDownloadNome + DateTime.Now.ToString("yyyy - MM - dd HH':'mm':'ss':'fff").Replace(" - ", "").Replace(":", "").Replace(" ", "") + ".xlsx";

                templateFileName = AdjustPath(HttpContext.Current.Server.MapPath("~")) + @"Arquivos\ExcelExportacao\Modelos\" + strDownloadNome + ".xlsx";

                GetNewFileMemory(ref templateFileName);

                if (File.Exists(fileName))
                    File.Delete(fileName);

                MemoryStream returnedMemoryStream = null;

                using (var file = new FileStream(fileName, FileMode.CreateNew))
                {
                    using (var temp = new FileStream(templateFileName, FileMode.Open))
                    {
                        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                        ExcelPackage package = new ExcelPackage(file, temp);

                        ExcelWorksheet wworksheet = package.Workbook.Worksheets["Data"];

                        wworksheet.View.ShowHeaders = true;
                        wworksheet.View.ShowGridLines = false;
                        //Titulo da planilha. Deixar linha 3 vazias formatadas no modelo excel.
                        wworksheet.Cells[3, 2].Value = tituloPlanilha;

                        int currentRow = 6;
                        int colNameTable = 3;
                        int lastColName = 0;


                        if (grafico != null)
                        {
                            // Preenche cabeçalho com os periodos:
                            //wworksheet.Cells[currentRow, colNameTable].Value = filtro.Onda1.FirstOrDefault().DescItem;
                            //colNameTable++;
                            //wworksheet.Cells[currentRow, colNameTable].Value = filtro.Onda2.FirstOrDefault().DescItem;
                            //colNameTable++;

                            //wworksheet.Cells[currentRow, colNameTable].Value = filtro.Onda3.FirstOrDefault().DescItem;
                            //colNameTable++;

                            //wworksheet.Cells[currentRow, colNameTable].Value = filtro.Onda4.FirstOrDefault().DescItem;
                            //colNameTable++;

                            if (filtro.Onda1.FirstOrDefault() != null)
                                wworksheet.Cells[currentRow, colNameTable].Value = filtro.Onda1.FirstOrDefault().DescItem;
                            else
                                wworksheet.Cells[currentRow, colNameTable].Value = "N/A";
                            colNameTable++;

                            if (filtro.Onda2.FirstOrDefault() != null)
                                wworksheet.Cells[currentRow, colNameTable].Value = filtro.Onda2.FirstOrDefault().DescItem;
                            else
                                wworksheet.Cells[currentRow, colNameTable].Value = "N/A";
                            colNameTable++;

                            if (filtro.Onda3.FirstOrDefault() != null)
                                wworksheet.Cells[currentRow, colNameTable].Value = filtro.Onda3.FirstOrDefault().DescItem;
                            else
                                wworksheet.Cells[currentRow, colNameTable].Value = "N/A";
                            colNameTable++;

                            if (filtro.Onda4.FirstOrDefault() != null)
                                wworksheet.Cells[currentRow, colNameTable].Value = filtro.Onda4.FirstOrDefault().DescItem;
                            else
                                wworksheet.Cells[currentRow, colNameTable].Value = "N/A";
                            colNameTable++;

                            //wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoColunas5.DescMarca;
                            //colNameTable++;

                            colNameTable--;
                            lastColName = colNameTable;

                            //Formata cabeçalho
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(3, 6) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Font.Size = 11;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(3, 6) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Font.Name = "Arial Nova Cond";
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(3, 6) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Font.Bold = true;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(3, 6) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(3, 6) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(3, 6) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(3, 6) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(3, 6) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(3, 6) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                            colNameTable = 2;
                            currentRow++;


                            ////Preenche lateral
                            //wworksheet.Cells[currentRow, colNameTable].Value = "Total Awareness";
                            //currentRow++;
                            //wworksheet.Cells[currentRow, colNameTable].Value = "Prompeted";
                            //currentRow++;
                            //wworksheet.Cells[currentRow, colNameTable].Value = "Spontaneous - OM";
                            //currentRow++;
                            //wworksheet.Cells[currentRow, colNameTable].Value = "Awareness - TOM";
                            //currentRow++;
                            //wworksheet.Cells[currentRow, colNameTable].Value = "Base";

                            //wworksheet.Cells[currentRow, colNameTable].Value = listaTraducaoComponente.FirstOrDefault(a => a.Objeto.Equals("dashboard-eight-leg-Awareness")).Texto;
                            //currentRow++;


                            wworksheet.Cells[currentRow, colNameTable].Value = listaTraducaoComponente.FirstOrDefault(a => a.Objeto.Equals("dashboard-eight-leg-Spontaneous")).Texto;
                            currentRow++;

                            wworksheet.Cells[currentRow, colNameTable].Value = listaTraducaoComponente.FirstOrDefault(a => a.Objeto.Equals("dashboard-eight-leg-Prompeted")).Texto;
                            currentRow++;


                            wworksheet.Cells[currentRow, colNameTable].Value = listaTraducaoComponente.FirstOrDefault(a => a.Objeto.Equals("dashboard-eight-leg-total-Awareness")).Texto;
                            currentRow++;

                            //wworksheet.Cells[currentRow, colNameTable].Value = listaTraducaoComponente.FirstOrDefault(a => a.Objeto.Equals("dashboard-eight-leg-outros")).Texto;
                            //currentRow++;
                            wworksheet.Cells[currentRow, colNameTable].Value = "Base";

                            currentRow--;
                            var lastRow = currentRow;
                            colNameTable = 3;
                            currentRow = 7;






                            //wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoColunas1.PercTotal;
                            //if (grafico.GraficoColunas1.TesteSigTotal == "MENOR")
                            //    wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Red);
                            //if (grafico.GraficoColunas1.TesteSigTotal == "MAIOR")
                            //    wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Blue);
                            //currentRow++;

                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoColunas1.PercPrompeted;
                            if (grafico.GraficoColunas1.TesteSIGPrompeted == "MENOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Red);
                            if (grafico.GraficoColunas1.TesteSIGPrompeted == "MAIOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Blue);
                            currentRow++;

                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoColunas1.PercOM;
                            if (grafico.GraficoColunas1.TesteSIGOM == "MENOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Red);
                            if (grafico.GraficoColunas1.TesteSIGOM == "MAIOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Blue);
                            currentRow++;

                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoColunas1.PercTOM;
                            if (grafico.GraficoColunas1.TesteSIGTOM == "MENOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Red);
                            if (grafico.GraficoColunas1.TesteSIGTOM == "MAIOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Blue);
                            currentRow++;

                            //wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoColunas1.PercOutros;
                            //if (grafico.GraficoColunas1.TesteSigOutros == "MENOR")
                            //    wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Red);
                            //if (grafico.GraficoColunas1.TesteSigOutros == "MAIOR")
                            //    wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Blue);
                            //currentRow++;

                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoColunas1.BaseAbs;
                            colNameTable++;
                            currentRow = 7;

                            //wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoColunas2.PercTotal;
                            //if (grafico.GraficoColunas2.TesteSigTotal == "MENOR")
                            //    wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Red);
                            //if (grafico.GraficoColunas2.TesteSigTotal == "MAIOR")
                            //    wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Blue);
                            //currentRow++;

                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoColunas2.PercPrompeted;
                            if (grafico.GraficoColunas2.TesteSIGPrompeted == "MENOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Red);
                            if (grafico.GraficoColunas2.TesteSIGPrompeted == "MAIOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Blue);
                            currentRow++;

                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoColunas2.PercOM;
                            if (grafico.GraficoColunas2.TesteSIGOM == "MENOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Red);
                            if (grafico.GraficoColunas2.TesteSIGOM == "MAIOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Blue);
                            currentRow++;

                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoColunas2.PercTOM;
                            if (grafico.GraficoColunas2.TesteSIGTOM == "MENOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Red);
                            if (grafico.GraficoColunas2.TesteSIGTOM == "MAIOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Blue);
                            currentRow++;

                            //wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoColunas2.PercOutros;
                            //if (grafico.GraficoColunas2.TesteSigOutros == "MENOR")
                            //    wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Red);
                            //if (grafico.GraficoColunas2.TesteSigOutros == "MAIOR")
                            //    wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Blue);
                            //currentRow++;

                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoColunas2.BaseAbs;
                            colNameTable++;
                            currentRow = 7;

                            //wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoColunas3.PercTotal;
                            //if (grafico.GraficoColunas3.TesteSigTotal == "MENOR")
                            //    wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Red);
                            //if (grafico.GraficoColunas3.TesteSigTotal == "MAIOR")
                            //    wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Blue);
                            //currentRow++;

                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoColunas3.PercPrompeted;
                            if (grafico.GraficoColunas3.TesteSIGPrompeted == "MENOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Red);
                            if (grafico.GraficoColunas3.TesteSIGPrompeted == "MAIOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Blue);
                            currentRow++;

                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoColunas3.PercOM;
                            if (grafico.GraficoColunas3.TesteSIGOM == "MENOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Red);
                            if (grafico.GraficoColunas3.TesteSIGOM == "MAIOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Blue);
                            currentRow++;

                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoColunas3.PercTOM;
                            if (grafico.GraficoColunas3.TesteSIGTOM == "MENOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Red);
                            if (grafico.GraficoColunas3.TesteSIGTOM == "MAIOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Blue);
                            currentRow++;

                            //wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoColunas3.PercOutros;
                            //if (grafico.GraficoColunas3.TesteSigOutros == "MENOR")
                            //    wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Red);
                            //if (grafico.GraficoColunas3.TesteSigOutros == "MAIOR")
                            //    wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Blue);
                            //currentRow++;

                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoColunas3.BaseAbs;
                            colNameTable++;
                            currentRow = 7;


                            //wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoColunas4.PercTotal;
                            //if (grafico.GraficoColunas4.TesteSigTotal == "MENOR")
                            //    wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Red);
                            //if (grafico.GraficoColunas4.TesteSigTotal == "MAIOR")
                            //    wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Blue);
                            //currentRow++;

                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoColunas4.PercPrompeted;
                            if (grafico.GraficoColunas4.TesteSIGPrompeted == "MENOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Red);
                            if (grafico.GraficoColunas4.TesteSIGPrompeted == "MAIOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Blue);
                            currentRow++;

                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoColunas4.PercOM;
                            if (grafico.GraficoColunas4.TesteSIGOM == "MENOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Red);
                            if (grafico.GraficoColunas4.TesteSIGOM == "MAIOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Blue);
                            currentRow++;

                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoColunas4.PercTOM;
                            if (grafico.GraficoColunas4.TesteSIGTOM == "MENOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Red);
                            if (grafico.GraficoColunas4.TesteSIGTOM == "MAIOR")
                                wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Blue);
                            currentRow++;

                            //wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoColunas4.PercOutros;
                            //if (grafico.GraficoColunas4.TesteSigOutros == "MENOR")
                            //    wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Red);
                            //if (grafico.GraficoColunas4.TesteSigOutros == "MAIOR")
                            //    wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Blue);
                            //currentRow++;

                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoColunas4.BaseAbs;
                            colNameTable++;
                            currentRow = 7;




                            //wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoColunas5.PercTotal;
                            //currentRow++;
                            //wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoColunas5.PercPrompeted;
                            //currentRow++;
                            //wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoColunas5.PercOM;
                            //currentRow++;
                            //wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoColunas5.PercTOM;
                            //currentRow++;
                            //wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoColunas5.BaseAbs;
                            //colNameTable++;
                            //currentRow = 7;


                            lastRow++;
                            currentRow = lastRow;
                            colNameTable--;
                            lastColName = colNameTable;

                            //Formata dados
                            //wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 7) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Numberformat.Format = "0.00";
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 7) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Font.Size = 11;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 7) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Font.Name = "Arial Nova Cond";
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 7) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Font.Bold = true;

                            // Preenche lateral
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 10) + ":" + ExcelUtil.GetCelulaByColumnIndex(3, 10)].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 10) + ":" + ExcelUtil.GetCelulaByColumnIndex(3, 10)].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 10) + ":" + ExcelUtil.GetCelulaByColumnIndex(3, 10)].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 10) + ":" + ExcelUtil.GetCelulaByColumnIndex(3, 10)].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 10) + ":" + ExcelUtil.GetCelulaByColumnIndex(3, 10)].Style.Border.Top.Style = ExcelBorderStyle.Thin;

                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 7) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 7) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 7) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 7) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 7) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                        }
                        wworksheet.Select();

                        using (MemoryStream memoryStream = new MemoryStream())
                        {
                            package.SaveAs(memoryStream);

                            memoryStream.Position = 0;

                            returnedMemoryStream = memoryStream;
                        }
                    }
                }

                return File.ReadAllBytes(@"" + fileName);
            }

            catch (Exception)
            {
                throw;
            }
            finally
            {
                File.Delete(fileName);
                File.Delete(templateFileName);
            }
        }
        #endregion


        #region DownloadDashboardNine

        public byte[] DownloadDashboardNineGraficoLinhasImagem(GraficoLinhasImagem grafico, string tituloPlanilha, List<TraducaoComponente> listaTraducaoComponente, FiltroPadraoExcel filtro)
        {
            string fileName = "";
            string templateFileName = "";
            try
            {
                string strDownloadNome;

                strDownloadNome = "Modelo_Exportação_Gráficos_Padrao_linhas";

                string strPath = AdjustPath(HttpContext.Current.Server.MapPath("~")) + @"Arquivos\ExcelExportacao\workspace\";
                fileName = strPath + strDownloadNome + DateTime.Now.ToString("yyyy - MM - dd HH':'mm':'ss':'fff").Replace(" - ", "").Replace(":", "").Replace(" ", "") + ".xlsx";

                templateFileName = AdjustPath(HttpContext.Current.Server.MapPath("~")) + @"Arquivos\ExcelExportacao\Modelos\" + strDownloadNome + ".xlsx";

                GetNewFileMemory(ref templateFileName);

                if (File.Exists(fileName))
                    File.Delete(fileName);

                MemoryStream returnedMemoryStream = null;

                using (var file = new FileStream(fileName, FileMode.CreateNew))
                {
                    using (var temp = new FileStream(templateFileName, FileMode.Open))
                    {
                        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                        ExcelPackage package = new ExcelPackage(file, temp);

                        ExcelWorksheet wworksheet = package.Workbook.Worksheets["Data"];

                        wworksheet.View.ShowHeaders = true;
                        wworksheet.View.ShowGridLines = false;
                        //Titulo da planilha. Deixar linha 3 vazias formatadas no modelo excel.
                        wworksheet.Cells[3, 2].Value = tituloPlanilha;

                        int currentRow = 6;
                        int colNameTable = 3;
                        int lastColName = 0;


                        if (grafico != null)
                        {

                            // Preenche cabeçalho com os periodos:



                            wworksheet.Cells[currentRow - 1, colNameTable].Value = filtro.Marca1.FirstOrDefault().DescItem;
                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoLinhasImagem1.FirstOrDefault().DescMarca;
                            colNameTable++;
                            wworksheet.Cells[currentRow, colNameTable].Value = "";
                            colNameTable++;

                            wworksheet.Cells[currentRow - 1, colNameTable].Value = filtro.Marca2.FirstOrDefault().DescItem;
                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoLinhasImagem2.FirstOrDefault().DescMarca;
                            colNameTable++;
                            wworksheet.Cells[currentRow, colNameTable].Value = "";
                            colNameTable++;

                            wworksheet.Cells[currentRow - 1, colNameTable].Value = filtro.Marca3.FirstOrDefault().DescItem;
                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoLinhasImagem3.FirstOrDefault().DescMarca;
                            colNameTable++;
                            wworksheet.Cells[currentRow, colNameTable].Value = "";
                            colNameTable++;

                            wworksheet.Cells[currentRow - 1, colNameTable].Value = filtro.Marca4.FirstOrDefault().DescItem;
                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoLinhasImagem4.FirstOrDefault().DescMarca;
                            colNameTable++;
                            wworksheet.Cells[currentRow, colNameTable].Value = "";
                            colNameTable++;

                            wworksheet.Cells[currentRow - 1, colNameTable].Value = filtro.Marca5.FirstOrDefault().DescItem;
                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoLinhasImagem5.FirstOrDefault().DescMarca;
                            colNameTable++;
                            wworksheet.Cells[currentRow, colNameTable].Value = "";
                            colNameTable++;


                            colNameTable = 3;


                            for (int i = 0; i < 5; i++)
                            {
                                wworksheet.Cells[currentRow, colNameTable].Value = "Anterior";
                                colNameTable++;
                                wworksheet.Cells[currentRow, colNameTable].Value = "Atual";
                                colNameTable++;

                            }


                            colNameTable--;
                            lastColName = colNameTable;

                            //Formata cabeçalho
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(3, 6) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Font.Size = 11;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(3, 6) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Font.Name = "Arial Nova Cond";
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(3, 6) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Font.Bold = true;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(3, 6) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(3, 6) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(3, 6) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(3, 6) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(3, 6) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(3, 6) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                            colNameTable = 2;
                            currentRow++;

                            ////Preenche lateral
                            foreach (var item in grafico.GraficoLinhasImagem1.Select(a => a.Descricao).Distinct().ToList())
                            {
                                wworksheet.Cells[currentRow, colNameTable].Value = item;
                                currentRow++;
                            }

                            wworksheet.Cells[currentRow, colNameTable].Value = "Base";
                            currentRow++;



                            currentRow--;
                            var lastRow = currentRow;
                            colNameTable = 3;
                            currentRow = 7;

                            foreach (var linha in grafico.GraficoLinhasImagem1)
                            {
                                colNameTable = 3;
                                wworksheet.Cells[currentRow, colNameTable].Value = linha.PercAnterior;
                                colNameTable++;
                                wworksheet.Cells[currentRow, colNameTable].Value = linha.PercAtual;
                                if (linha.TestSig == "MENOR")
                                    wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Red);
                                if (linha.TestSig == "MAIOR")
                                    wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Blue);
                                colNameTable++;
                                currentRow++;
                            }

                            colNameTable = 3;
                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoLinhasImagem1[0].BaseAnt;
                            colNameTable++;
                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoLinhasImagem1[0].Base;

                            colNameTable++;
                            currentRow++;


                            currentRow = 7;

                            foreach (var linha in grafico.GraficoLinhasImagem2)
                            {
                                colNameTable = 5;
                                wworksheet.Cells[currentRow, colNameTable].Value = linha.PercAnterior;
                                colNameTable++;
                                wworksheet.Cells[currentRow, colNameTable].Value = linha.PercAtual;
                                if (linha.TestSig == "MENOR")
                                    wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Red);
                                if (linha.TestSig == "MAIOR")
                                    wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Blue);
                                colNameTable++;
                                currentRow++;
                            }

                            colNameTable = 5;
                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoLinhasImagem2[0].BaseAnt;
                            colNameTable++;
                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoLinhasImagem2[0].Base;

                            colNameTable++;
                            currentRow++;

                            currentRow = 7;
                            foreach (var linha in grafico.GraficoLinhasImagem3)
                            {
                                colNameTable = 7;
                                wworksheet.Cells[currentRow, colNameTable].Value = linha.PercAnterior;
                                colNameTable++;
                                wworksheet.Cells[currentRow, colNameTable].Value = linha.PercAtual;
                                if (linha.TestSig == "MENOR")
                                    wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Red);
                                if (linha.TestSig == "MAIOR")
                                    wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Blue);
                                colNameTable++;
                                currentRow++;
                            }
                            colNameTable = 7;
                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoLinhasImagem3[0].BaseAnt;
                            colNameTable++;
                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoLinhasImagem3[0].Base;

                            colNameTable++;
                            currentRow++;


                            currentRow = 7;
                            foreach (var linha in grafico.GraficoLinhasImagem4)
                            {
                                colNameTable = 9;
                                wworksheet.Cells[currentRow, colNameTable].Value = linha.PercAnterior;
                                colNameTable++;
                                wworksheet.Cells[currentRow, colNameTable].Value = linha.PercAtual;
                                if (linha.TestSig == "MENOR")
                                    wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Red);
                                if (linha.TestSig == "MAIOR")
                                    wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Blue);
                                colNameTable++;
                                currentRow++;
                            }
                            colNameTable = 9;
                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoLinhasImagem4[0].BaseAnt;
                            colNameTable++;
                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoLinhasImagem4[0].Base;

                            colNameTable++;
                            currentRow++;


                            currentRow = 7;
                            foreach (var linha in grafico.GraficoLinhasImagem5)
                            {
                                colNameTable = 11;
                                wworksheet.Cells[currentRow, colNameTable].Value = linha.PercAnterior;
                                colNameTable++;
                                wworksheet.Cells[currentRow, colNameTable].Value = linha.PercAtual;
                                if (linha.TestSig == "MENOR")
                                    wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Red);
                                if (linha.TestSig == "MAIOR")
                                    wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Blue);
                                colNameTable++;
                                currentRow++;
                            }

                            colNameTable = 11;
                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoLinhasImagem5[0].BaseAnt;
                            colNameTable++;
                            wworksheet.Cells[currentRow, colNameTable].Value = grafico.GraficoLinhasImagem5[0].Base;

                            colNameTable++;
                            currentRow++;




                            //foreach (var item in grafico.Select(a => a.Descricao).Distinct().ToList())
                            //{
                            //    colNameTable = 3;
                            //    var itemLinha = grafico.Where(a => a.Descricao == item).ToList();

                            //    // MONTA COLUNAS DADOS
                            //    foreach (var dadoLinha in itemLinha)
                            //    {
                            //        wworksheet.Cells[currentRow, colNameTable].Value = dadoLinha.Perc;
                            //        colNameTable++;
                            //    }

                            //    currentRow++;
                            //}

                            currentRow = 7;

                            lastRow++;
                            currentRow = lastRow;
                            colNameTable--;
                            lastColName = colNameTable;

                            //Formata dados
                            //wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 7) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Numberformat.Format = "0.0";
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 7) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Font.Size = 11;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 7) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Font.Name = "Arial Nova Cond";
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 7) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Font.Bold = true;

                            // Preenche lateral
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 11) + ":" + ExcelUtil.GetCelulaByColumnIndex(3, 11)].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 11) + ":" + ExcelUtil.GetCelulaByColumnIndex(3, 11)].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 11) + ":" + ExcelUtil.GetCelulaByColumnIndex(3, 11)].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 11) + ":" + ExcelUtil.GetCelulaByColumnIndex(3, 11)].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 11) + ":" + ExcelUtil.GetCelulaByColumnIndex(3, 11)].Style.Border.Top.Style = ExcelBorderStyle.Thin;

                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 7) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 7) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 7) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 7) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 7) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                        }
                        wworksheet.Select();

                        using (MemoryStream memoryStream = new MemoryStream())
                        {
                            package.SaveAs(memoryStream);

                            memoryStream.Position = 0;

                            returnedMemoryStream = memoryStream;
                        }
                    }
                }

                return File.ReadAllBytes(@"" + fileName);
            }

            catch (Exception)
            {
                throw;
            }
            finally
            {
                File.Delete(fileName);
                File.Delete(templateFileName);
            }
        }


        public byte[] DownloadDashboardNineImagemPura(List<GraficoImagemEvolutivoProcedure> grafico, string tituloPlanilha, List<TraducaoComponente> listaTraducaoComponente)
        {
            string fileName = "";
            string templateFileName = "";
            try
            {
                string strDownloadNome;

                strDownloadNome = "Modelo_Exportação_Gráficos_Padrao_Four";

                string strPath = AdjustPath(HttpContext.Current.Server.MapPath("~")) + @"Arquivos\ExcelExportacao\workspace\";
                fileName = strPath + strDownloadNome + DateTime.Now.ToString("yyyy - MM - dd HH':'mm':'ss':'fff").Replace(" - ", "").Replace(":", "").Replace(" ", "") + ".xlsx";

                templateFileName = AdjustPath(HttpContext.Current.Server.MapPath("~")) + @"Arquivos\ExcelExportacao\Modelos\" + strDownloadNome + ".xlsx";

                GetNewFileMemory(ref templateFileName);

                if (File.Exists(fileName))
                    File.Delete(fileName);

                MemoryStream returnedMemoryStream = null;

                using (var file = new FileStream(fileName, FileMode.CreateNew))
                {
                    using (var temp = new FileStream(templateFileName, FileMode.Open))
                    {
                        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                        ExcelPackage package = new ExcelPackage(file, temp);

                        ExcelWorksheet wworksheet = package.Workbook.Worksheets["Data"];

                        wworksheet.View.ShowHeaders = true;
                        wworksheet.View.ShowGridLines = false;
                        //Titulo da planilha. Deixar linha 3 vazias formatadas no modelo excel.
                        wworksheet.Cells[3, 2].Value = tituloPlanilha;

                        int currentRow = 6;
                        int colNameTable = 3;
                        int lastColName = 0;


                        if (grafico != null)
                        {

                            // Preenche cabeçalho com os periodos:
                            foreach (var item in grafico.Select(a => a.DescMarca).Distinct().ToList())
                            {
                                wworksheet.Cells[currentRow, colNameTable].Value = item;
                                colNameTable++;
                                wworksheet.Cells[currentRow, colNameTable].Value = "Base";
                                colNameTable++;
                            }

                            colNameTable--;
                            lastColName = colNameTable;

                            //Formata cabeçalho
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(3, 6) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Font.Size = 11;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(3, 6) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Font.Name = "Arial Nova Cond";
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(3, 6) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Font.Bold = true;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(3, 6) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(3, 6) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(3, 6) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(3, 6) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(3, 6) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(3, 6) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                            colNameTable = 2;
                            currentRow++;

                            ////Preenche lateral
                            foreach (var item in grafico.Select(a => a.DescOnda).Distinct().ToList())
                            {
                                wworksheet.Cells[currentRow, colNameTable].Value = item;
                                currentRow++;
                            }

                            currentRow--;
                            var lastRow = currentRow;
                            colNameTable = 3;
                            currentRow = 7;

                            foreach (var item in grafico.Select(a => a.DescOnda).Distinct().ToList())
                            {
                                colNameTable = 3;
                                var itemLinha = grafico.Where(a => a.DescOnda == item).ToList();

                                // MONTA COLUNAS DADOS
                                foreach (var dadoLinha in itemLinha)
                                {
                                    wworksheet.Cells[currentRow, colNameTable].Value = dadoLinha.Perc;
                                    if (dadoLinha.TesteSig == "MENOR")
                                        wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Red);
                                    if (dadoLinha.TesteSig == "MAIOR")
                                        wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Blue);
                                    colNameTable++;

                                    wworksheet.Cells[currentRow, colNameTable].Value = dadoLinha.BaseAbs;
                                    colNameTable++;
                                }

                                currentRow++;
                            }

                            //currentRow = 7;

                            //lastRow++;
                            //currentRow = lastRow;
                            colNameTable--;
                            //lastColName = colNameTable;

                            //Formata dados
                            //wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 7) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Numberformat.Format = "0.00";
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 7) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Font.Size = 11;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 7) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Font.Name = "Arial Nova Cond";
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 7) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Font.Bold = true;

                            // Preenche lateral
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 11) + ":" + ExcelUtil.GetCelulaByColumnIndex(3, 11)].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 11) + ":" + ExcelUtil.GetCelulaByColumnIndex(3, 11)].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 11) + ":" + ExcelUtil.GetCelulaByColumnIndex(3, 11)].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 11) + ":" + ExcelUtil.GetCelulaByColumnIndex(3, 11)].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 11) + ":" + ExcelUtil.GetCelulaByColumnIndex(3, 11)].Style.Border.Top.Style = ExcelBorderStyle.Thin;

                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 7) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 7) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 7) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 7) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 7) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                        }
                        wworksheet.Select();

                        using (MemoryStream memoryStream = new MemoryStream())
                        {
                            package.SaveAs(memoryStream);

                            memoryStream.Position = 0;

                            returnedMemoryStream = memoryStream;
                        }
                    }
                }

                return File.ReadAllBytes(@"" + fileName);
            }

            catch (Exception)
            {
                throw;
            }
            finally
            {
                File.Delete(fileName);
                File.Delete(templateFileName);
            }
        }

        public byte[] DownloadDashboardNineImagemPura2(List<GraficoImagemEvolutivoProcedure> grafico, string tituloPlanilha, List<TraducaoComponente> listaTraducaoComponente)
        {
            string fileName = "";
            string templateFileName = "";
            try
            {
                string strDownloadNome;

                strDownloadNome = "Modelo_Exportação_Gráficos_Padrao_Four";

                string strPath = AdjustPath(HttpContext.Current.Server.MapPath("~")) + @"Arquivos\ExcelExportacao\workspace\";
                fileName = strPath + strDownloadNome + DateTime.Now.ToString("yyyy - MM - dd HH':'mm':'ss':'fff").Replace(" - ", "").Replace(":", "").Replace(" ", "") + ".xlsx";

                templateFileName = AdjustPath(HttpContext.Current.Server.MapPath("~")) + @"Arquivos\ExcelExportacao\Modelos\" + strDownloadNome + ".xlsx";

                GetNewFileMemory(ref templateFileName);

                if (File.Exists(fileName))
                    File.Delete(fileName);

                MemoryStream returnedMemoryStream = null;

                using (var file = new FileStream(fileName, FileMode.CreateNew))
                {
                    using (var temp = new FileStream(templateFileName, FileMode.Open))
                    {
                        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                        ExcelPackage package = new ExcelPackage(file, temp);

                        ExcelWorksheet wworksheet = package.Workbook.Worksheets["Data"];

                        wworksheet.View.ShowHeaders = true;
                        wworksheet.View.ShowGridLines = false;
                        //Titulo da planilha. Deixar linha 3 vazias formatadas no modelo excel.
                        wworksheet.Cells[3, 2].Value = tituloPlanilha;

                        int currentRow = 6;
                        int colNameTable = 3;
                        int lastColName = 0;


                        if (grafico != null)
                        {

                            // Preenche cabeçalho com os periodos:
                            foreach (var item in grafico.Select(a => a.DescMarca).Distinct().ToList())
                            {
                                wworksheet.Cells[currentRow, colNameTable].Value = item;
                                colNameTable++;
                                wworksheet.Cells[currentRow, colNameTable].Value = "Base";
                                colNameTable++;
                            }

                            colNameTable--;
                            lastColName = colNameTable;

                            //Formata cabeçalho
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(3, 6) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Font.Size = 11;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(3, 6) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Font.Name = "Arial Nova Cond";
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(3, 6) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Font.Bold = true;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(3, 6) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(3, 6) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(3, 6) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(3, 6) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(3, 6) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(3, 6) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                            colNameTable = 2;
                            currentRow++;

                            ////Preenche lateral
                            foreach (var item in grafico.Select(a => a.DescOnda).Distinct().ToList())
                            {
                                wworksheet.Cells[currentRow, colNameTable].Value = item;
                                currentRow++;
                            }

                            currentRow--;
                            var lastRow = currentRow;
                            colNameTable = 3;
                            currentRow = 7;

                            foreach (var item in grafico.Select(a => a.DescOnda).Distinct().ToList())
                            {
                                colNameTable = 3;
                                var itemLinha = grafico.Where(a => a.DescOnda == item).ToList();

                                // MONTA COLUNAS DADOS
                                foreach (var dadoLinha in itemLinha)
                                {
                                    wworksheet.Cells[currentRow, colNameTable].Value = dadoLinha.Perc;
                                    if (dadoLinha.TesteSig == "MENOR")
                                        wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Red);
                                    if (dadoLinha.TesteSig == "MAIOR")
                                        wworksheet.Cells[currentRow, colNameTable].Style.Font.Color.SetColor(Color.Blue);
                                    colNameTable++;

                                    wworksheet.Cells[currentRow, colNameTable].Value = dadoLinha.BaseAbs;
                                    colNameTable++;
                                }

                                currentRow++;
                            }

                            //currentRow = 7;

                            //lastRow++;
                            //currentRow = lastRow;
                            colNameTable--;
                            //lastColName = colNameTable;

                            //Formata dados
                            //wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 7) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Numberformat.Format = "0.00";
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 7) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Font.Size = 11;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 7) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Font.Name = "Arial Nova Cond";
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 7) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Font.Bold = true;

                            // Preenche lateral
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 11) + ":" + ExcelUtil.GetCelulaByColumnIndex(3, 11)].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 11) + ":" + ExcelUtil.GetCelulaByColumnIndex(3, 11)].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 11) + ":" + ExcelUtil.GetCelulaByColumnIndex(3, 11)].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 11) + ":" + ExcelUtil.GetCelulaByColumnIndex(3, 11)].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 11) + ":" + ExcelUtil.GetCelulaByColumnIndex(3, 11)].Style.Border.Top.Style = ExcelBorderStyle.Thin;

                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 7) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 7) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 7) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 7) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 7) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                        }
                        wworksheet.Select();

                        using (MemoryStream memoryStream = new MemoryStream())
                        {
                            package.SaveAs(memoryStream);

                            memoryStream.Position = 0;

                            returnedMemoryStream = memoryStream;
                        }
                    }
                }

                return File.ReadAllBytes(@"" + fileName);
            }

            catch (Exception)
            {
                throw;
            }
            finally
            {
                File.Delete(fileName);
                File.Delete(templateFileName);
            }
        }



        public byte[] DownloadDashboardNineTabelaAdHocAtributo(TabelaPadraoAdHoc grafico, string tituloPlanilha, List<TraducaoComponente> listaTraducaoComponente, PadraoComboFiltro marca)
        {
            string fileName = "";
            string templateFileName = "";
            try
            {
                string strDownloadNome;

                strDownloadNome = "Modelo_Exportação_Gráficos_Padrao_Four";

                string strPath = AdjustPath(HttpContext.Current.Server.MapPath("~")) + @"Arquivos\ExcelExportacao\workspace\";
                fileName = strPath + strDownloadNome + DateTime.Now.ToString("yyyy - MM - dd HH':'mm':'ss':'fff").Replace(" - ", "").Replace(":", "").Replace(" ", "") + ".xlsx";

                templateFileName = AdjustPath(HttpContext.Current.Server.MapPath("~")) + @"Arquivos\ExcelExportacao\Modelos\" + strDownloadNome + ".xlsx";

                GetNewFileMemory(ref templateFileName);

                if (File.Exists(fileName))
                    File.Delete(fileName);

                MemoryStream returnedMemoryStream = null;

                using (var file = new FileStream(fileName, FileMode.CreateNew))
                {
                    using (var temp = new FileStream(templateFileName, FileMode.Open))
                    {
                        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                        ExcelPackage package = new ExcelPackage(file, temp);

                        ExcelWorksheet wworksheet = package.Workbook.Worksheets["Data"];

                        wworksheet.View.ShowHeaders = true;
                        wworksheet.View.ShowGridLines = false;
                        //Titulo da planilha. Deixar linha 3 vazias formatadas no modelo excel.
                        wworksheet.Cells[3, 2].Value = tituloPlanilha;

                        int currentRow = 6;
                        int colNameTable = 3;
                        int lastColName = 0;


                        if (grafico != null)
                        {

                            if (marca != null)
                            {
                                Color colour = ColorTranslator.FromHtml(marca.CorItem);
                                wworksheet.Cells[currentRow, colNameTable - 1].Style.Font.Color.SetColor(colour);
                                wworksheet.Cells[currentRow, colNameTable - 1].Value = marca.DescItem;
                            }
                            else
                            {
                                wworksheet.Cells[currentRow, colNameTable - 1].Value = "";
                            }


                            wworksheet.Cells[currentRow, colNameTable - 1].Style.Font.Bold = true;
                            wworksheet.Cells[currentRow, colNameTable - 1].Style.Font.Size = 14;

                            // Preenche cabeçalho com os periodos:
                            //foreach (var item in grafico.Titulos.Where(x => x.DescAtributo.Length > 5).ToList())
                            foreach (var item in grafico.Titulos.ToList())
                            {
                                wworksheet.Cells[currentRow, colNameTable].Value = item.DescAtributo;
                                colNameTable++;


                            }
                            wworksheet.Row(currentRow).Height = 42;

                            colNameTable--;
                            lastColName = colNameTable;

                            //Formata cabeçalho
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(3, 6) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Font.Size = 11;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(3, 6) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Font.Name = "Arial Nova Cond";
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(3, 6) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Font.Bold = true;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(3, 6) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(3, 6) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(3, 6) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(3, 6) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(3, 6) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(3, 6) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                            colNameTable = 2;
                            currentRow++;

                            ////Preenche lateral
                            foreach (var item in grafico.Dados)
                            {
                                wworksheet.Cells[currentRow, colNameTable].Value = item.Texto;
                                currentRow++;
                            }

                            currentRow--;
                            var lastRow = currentRow;
                            colNameTable = 3;
                            currentRow = 7;


                            colNameTable = 3;

                            //var qtdColunas = grafico.Titulos.Where(x => x.DescAtributo.Length > 5).Count();
                            var qtdColunas = grafico.Titulos.Count();

                            // MONTA COLUNAS DADOS
                            foreach (var dadoLinha in grafico.Dados)
                            {

                                wworksheet.Cells[currentRow, colNameTable].Value = dadoLinha.Perc1;
                                colNameTable++;

                                wworksheet.Cells[currentRow, colNameTable].Value = dadoLinha.Perc2;
                                colNameTable++;

                                wworksheet.Cells[currentRow, colNameTable].Value = dadoLinha.Perc3;
                                colNameTable++;

                                wworksheet.Cells[currentRow, colNameTable].Value = dadoLinha.Perc4;
                                colNameTable++;

                                if (qtdColunas >= 5)
                                {
                                    wworksheet.Cells[currentRow, colNameTable].Value = dadoLinha.Perc5;
                                    colNameTable++;
                                }

                                if (qtdColunas >= 6)
                                {
                                    wworksheet.Cells[currentRow, colNameTable].Value = dadoLinha.Perc6;
                                    colNameTable++;
                                }

                                if (qtdColunas >= 7)
                                {
                                    wworksheet.Cells[currentRow, colNameTable].Value = dadoLinha.Perc7;
                                    colNameTable++;
                                }

                                if (qtdColunas >= 8)
                                {
                                    wworksheet.Cells[currentRow, colNameTable].Value = dadoLinha.Perc8;
                                    colNameTable++;
                                }

                                if (qtdColunas >= 9)
                                {
                                    wworksheet.Cells[currentRow, colNameTable].Value = dadoLinha.Perc9;
                                    colNameTable++;
                                }

                                if (qtdColunas >= 10)
                                {
                                    wworksheet.Cells[currentRow, colNameTable].Value = dadoLinha.Perc10;
                                    colNameTable++;
                                }


                                colNameTable = 3;
                                currentRow++;
                            }




                            //currentRow = 7;

                            //lastRow++;
                            //currentRow = lastRow;
                            colNameTable--;
                            //lastColName = colNameTable;

                            //Formata dados
                            //wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 7) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Numberformat.Format = "0.00";
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 7) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Font.Size = 11;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 7) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Font.Name = "Arial Nova Cond";
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 7) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Font.Bold = true;

                            // Preenche lateral
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 11) + ":" + ExcelUtil.GetCelulaByColumnIndex(3, 11)].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 11) + ":" + ExcelUtil.GetCelulaByColumnIndex(3, 11)].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 11) + ":" + ExcelUtil.GetCelulaByColumnIndex(3, 11)].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 11) + ":" + ExcelUtil.GetCelulaByColumnIndex(3, 11)].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 11) + ":" + ExcelUtil.GetCelulaByColumnIndex(3, 11)].Style.Border.Top.Style = ExcelBorderStyle.Thin;

                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 7) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 7) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 7) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 7) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            wworksheet.Cells[ExcelUtil.GetCelulaByColumnIndex(2, 7) + ":" + ExcelUtil.GetCelulaByColumnIndex(lastColName, currentRow)].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                        }
                        wworksheet.Select();

                        using (MemoryStream memoryStream = new MemoryStream())
                        {
                            package.SaveAs(memoryStream);

                            memoryStream.Position = 0;

                            returnedMemoryStream = memoryStream;
                        }
                    }
                }

                return File.ReadAllBytes(@"" + fileName);
            }

            catch (Exception ex)
            {
                throw;
            }
            finally
            {
                File.Delete(fileName);
                File.Delete(templateFileName);
            }
        }

        #endregion

    }
}
