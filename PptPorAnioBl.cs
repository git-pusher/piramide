using Intercase.CaseFramework.DataTransfer;
using AnalisisSiniestrosDanos.BusinessLogic.Base;
using AnalisisSiniestrosDanos.BusinessLogic.PPT.Base;
using AnalisisSiniestrosDanos.BusinessLogic.Reporte;
using AnalisisSiniestrosDanos.BusinessLogic;
using AnalisisSiniestrosDanos.Entities.Reporte;
using System;
using System.Collections.Generic;
using System.Linq;
using AnalisisSiniestrosDanos.BusinessEntities.Reporte;
using AnalisisSiniestrosDanos.Entities;

namespace AnalisisSiniestrosDanos.BusinessLogic.PPT
{
    public class PptPorAnioBl : INetOfficeTablaPpt, INetOfficeGraficaPpt
    {

        #region propiedades de solo lectura

        private readonly decimal IdCuenta = 0;
        private readonly decimal IdReporte = 0;

        #endregion propiedades de solo lectura

        #region constructores

        public PptPorAnioBl(decimal idCuenta, decimal idReporte)
        {
            IdCuenta = idCuenta;
            IdReporte = idReporte;         
        }

        #endregion constructores

        #region ITablaPpt

        public NetOffice.PowerPointApi._Slide AgregarTablaEnDiapositiva(NetOffice.PowerPointApi._Slide diapositiva)
        {
            IEnumerable<DatosTablaDto> datos = (IEnumerable<DatosTablaDto>)RecuperarDatosParaTabla();

            NetOffice.PowerPointApi.Table tabla = null;
            tabla = diapositiva.Shapes[2].Table;

            IEnumerable<String> titulos = datos.First().DatoTablaDtosList.Select(x => x.Titulo).Distinct();
            IEnumerable<String> conceptos = datos.First().DatoTablaDtosList.Select(x => x.Concepto).Distinct();

            int cTitulos = 1;
            foreach (String titulo in titulos)
            {
                if (cTitulos != 1)
                {
                    tabla.Columns.Add();
                    tabla.Columns.Add();
                    tabla.Columns[cTitulos * 2].Cells[1].Merge(tabla.Columns[(cTitulos * 2)+1].Cells[1]);
                }
                tabla.Columns[cTitulos*2].Cells[1].Shape.TextFrame.TextRange.Text = titulo;
                tabla.Columns[cTitulos * 2].Cells[2].Shape.TextFrame.TextRange.Text = "#";
                tabla.Columns[(cTitulos * 2) + 1].Cells[2].Shape.TextFrame.TextRange.Text = "$";
                cTitulos++;
            }

            int filas = 2;
            bool columnaObscura = true;
            foreach (String concepto in conceptos)
            {
                tabla.Rows.Add(filas + 1);
                columnaObscura = !columnaObscura;
                filas++;
                cTitulos = 1;
                if (concepto.Equals(ConstantesBl.TOTAL_GENERAL))
                {
                    tabla.Rows[filas + 1].Delete();
                    tabla.Rows[filas].Delete();
                }
                tabla.Rows[filas].Cells[1].Shape.TextFrame.TextRange.Text = concepto;
                foreach (String titulo in titulos)
                {
                    if (columnaObscura)
                    {
                        tabla.Cell(filas, cTitulos * 2).Shape.Fill.ForeColor.RGB = tabla.Cell(2, 2).Shape.Fill.ForeColor.RGB;
                        tabla.Cell(filas, (cTitulos * 2) + 1).Shape.Fill.ForeColor.RGB = tabla.Cell(2, 2).Shape.Fill.ForeColor.RGB;
                    }
                    else
                    {
                        tabla.Cell(filas, cTitulos * 2).Shape.Fill.ForeColor.RGB = tabla.Cell(3, 2).Shape.Fill.ForeColor.RGB;
                        tabla.Cell(filas, (cTitulos * 2) + 1).Shape.Fill.ForeColor.RGB = tabla.Cell(3, 2).Shape.Fill.ForeColor.RGB;
                    }
                    try
                    {
                        tabla.Rows[filas].Cells[cTitulos * 2].Shape.TextFrame.TextRange.Text = datos.First().DatoTablaDtosList.Where(x => x.Concepto == concepto && x.Titulo == titulo).Select(x => x.Cantidad).First().ToString();
                    }
                    catch
                    {
                        tabla.Rows[filas].Cells[cTitulos * 2].Shape.TextFrame.TextRange.Text = "-";
                    }
                    try
                    {
                        tabla.Rows[filas].Cells[(cTitulos * 2) + 1].Shape.TextFrame.TextRange.Text = datos.First().DatoTablaDtosList.Where(x => x.Concepto == concepto && x.Titulo == titulo).Select(x => x.Importe).First().ToString("#,##0.00");
                    }
                    catch
                    {
                        tabla.Rows[filas].Cells[(cTitulos * 2) + 1].Shape.TextFrame.TextRange.Text = "-";
                    }
                    cTitulos++;
                }
            }

            NetOffice.PowerPointApi.Shape shape = null;
            //Busca el objeto tabla en el indice indicado
            shape = diapositiva.Shapes[5];
            SeccionesReporteBl seccionesReporteBl = new SeccionesReporteBl();
            IEnumerable<SeccionesReporteDto> reportePolizaSeccionDtos = seccionesReporteBl.ObtieneSeccionesPorReporte(IdReporte);
            if (reportePolizaSeccionDtos.Where(x => x.CodSecRep.Equals(ConstantesBl.CODIGO_SECCION_REPORTE_ANALISIS_POR_ANIO)).Count() > 0 && !String.IsNullOrEmpty(reportePolizaSeccionDtos.Where(x => x.CodSecRep.Equals(ConstantesBl.CODIGO_SECCION_REPORTE_ANALISIS_POR_ANIO)).First().TxComentarios))
            {
                shape.TextFrame.TextRange.Text = "Comentarios: " + reportePolizaSeccionDtos.Where(x => x.CodSecRep.Equals(ConstantesBl.CODIGO_SECCION_REPORTE_ANALISIS_POR_ANIO)).First().TxComentarios;
            }

            return diapositiva;
        }

        public IList<int> NumElemTabla
        {
            get { return new List<int>() { 4 }; }
        }

        public IEnumerable<DataTransferObject> RecuperarDatosParaTabla()
        {
            PorAnioBl porAnioBl = new PorAnioBl();
            DatosTablaDto datosTablaDto = porAnioBl.ObtenerDatosDeSeccion(IdCuenta, IdReporte);
            List<DatosTablaDto> datosTablaDtoList = new List<DatosTablaDto>();
            datosTablaDtoList.Add(datosTablaDto);
            IEnumerable<DatosTablaDto> IDatosTablaDto = datosTablaDtoList;
            return IDatosTablaDto;
        }

        public string CodUsuario
        {
            get;
            set;
        }

        #endregion ITablaPpt

        #region IGraficaPpt

        public NetOffice.PowerPointApi._Slide AgregarGraficaEnDiapositiva(NetOffice.PowerPointApi._Slide diapositiva)
        {
            diapositiva = AgregarGraficaPorAnio(diapositiva);
            diapositiva = AgregarGraficaPorAnioGeneral(diapositiva);

            return diapositiva;
        }

        public IList<int> NumElemGrafica
        {
            get
            {
                return new List<int>()
                {
                   4, 5, 6 
                };
            }
        }

        #endregion IGraficaPpt

        #region Metodos Privados

        private NetOffice.PowerPointApi._Slide AgregarGraficaPorAnioGeneral(NetOffice.PowerPointApi._Slide diapositiva)
        {
            IEnumerable<DatosTablaDto> datos = (IEnumerable<DatosTablaDto>)RecuperarDatosParaTabla();

            if (datos.Select(x => x.DatoTablaDtosList).Distinct().Count() > 1)
            {
                return this.AgregarGraficaDatosAnioGeneral(diapositiva);
            }

            return this.AgregarGraficaDatosAnioGeneral(diapositiva);

        }

        private NetOffice.PowerPointApi._Slide AgregarGraficaPorAnio(NetOffice.PowerPointApi._Slide diapositiva)
        {
            IEnumerable<DatosTablaDto> datos = (IEnumerable<DatosTablaDto>)RecuperarDatosParaTabla();

            if (datos.Select(x => x.DatoTablaDtosList).Distinct().Count() > 1)
            {
                return this.AgregarGraficaDatosAnio(diapositiva);
            }

            return this.AgregarGraficaDatosAnio(diapositiva);
        }

        private NetOffice.PowerPointApi._Slide AgregarGraficaDatosAnio(NetOffice.PowerPointApi._Slide diapositiva)
        {
            NetOffice.ExcelApi._Workbook libro = null;
            NetOffice.ExcelApi._Worksheet hoja = null;
            NetOffice.PowerPointApi.ChartData graficaInfo = null;
            NetOffice.PowerPointApi.Chart grafica = null;


            try
            {
                libro = (NetOffice.ExcelApi._Workbook)graficaInfo.Workbook;
                hoja = (NetOffice.ExcelApi.Worksheet)libro.Worksheets[8];

                grafica = (NetOffice.PowerPointApi.Chart)diapositiva.Shapes[NumElemGrafica[1]].Chart;
                graficaInfo = grafica.ChartData;

                IEnumerable<DatosTablaDto> datos = (IEnumerable<DatosTablaDto>)RecuperarDatosParaTabla();
                IEnumerable<DatoTablaDto> datosGrafica = datos.First().DatoTablaDtosList;

                IEnumerable<String> anios = datosGrafica.Where(x => !(x.Titulo.Equals(ConstantesBl.TOTAL))).Select(x => x.Titulo).Distinct();
                IEnumerable<String> tipos = datosGrafica.Where(x => !(x.Concepto.Equals(ConstantesBl.TOTAL_GENERAL))).Select(x => x.Concepto).Distinct();

                NetOffice.ExcelApi.Range rango = (NetOffice.ExcelApi.Range)hoja.Cells.get_Range("A1", "B4");
               NetOffice.ExcelApi.ListObject tabla = hoja.ListObjects[1];

                tabla.Resize(rango);

                int numRenglon = 2;
                int cAnios = 0;
                String columna = "";
                if (datosGrafica != null && datosGrafica.Count() > 0)
                {
                    numRenglon = 3;
                    foreach (String tipo in tipos)
                    {
                        ((NetOffice.ExcelApi.Range)hoja.Cells.get_Range(string.Format("A{0}", numRenglon), Type.Missing)).FormulaR1C1 = TipoDatoBase.FormatearDato(TipoDatoAtrib.TEXTO, tipo);
                        columna = "";
                        int cColumna = 1;
                        cAnios = 0;
                        foreach (String anio in anios)
                        {
                            if (cAnios >= 2 || anios.Count() <= 3)
                            {
                                if (cColumna == 1)
                                    columna = "B{0}";
                                else if (cColumna == 2)
                                    columna = "C{0}";
                                else if (cColumna == 3)
                                    columna = "D{0}";
                                else if (cColumna == 4)
                                    columna = "E{0}";
                                else if (cColumna == 5)
                                    columna = "F{0}";
                                if (numRenglon == 3)
                                    ((NetOffice.ExcelApi.Range)hoja.Cells.get_Range(string.Format(columna, 2), Type.Missing)).FormulaR1C1 = TipoDatoBase.FormatearDato(TipoDatoAtrib.TEXTO, anio);
                                try
                                {
                                    ((NetOffice.ExcelApi.Range)hoja.Cells.get_Range(string.Format(columna, numRenglon), Type.Missing)).FormulaR1C1 = TipoDatoBase.FormatearDato(TipoDatoAtrib.ENTERO, datosGrafica.Where(x => x.Titulo.Equals(anio) && x.Concepto.Equals(tipo)).Select(x => x.Importe).First());
                                }
                                catch
                                {
                                    ((NetOffice.ExcelApi.Range)hoja.Cells.get_Range(string.Format(columna, numRenglon), Type.Missing)).FormulaR1C1 = 0;
                                }
                                cColumna++;
                            }
                            cAnios++;
                        }
                        numRenglon++;
                    }

                    grafica.ApplyDataLabels(NetOffice.PowerPointApi.Enums.XlDataLabelsType.xlDataLabelsShowNone);
                }
                grafica.Refresh();    

            }
            catch (Exception ex)
            {
                Excepcion.InsertarException(ex, CodUsuario);
            }
            finally
            {
                if (libro != null)
                {
                    libro.Save();
                    libro.Close(true);
                    grafica.Application.Quit();
                    grafica.Dispose();
                }
            }

            return diapositiva;
        }

        private NetOffice.PowerPointApi._Slide AgregarGraficaDatosAnioGeneral(NetOffice.PowerPointApi._Slide diapositiva)
        {        
            NetOffice.PowerPointApi.Chart grafica = null;
            NetOffice.PowerPointApi.ChartData graficaInfo = null;
            NetOffice.ExcelApi._Workbook cuaderno = null;
            NetOffice.ExcelApi._Worksheet hoja = null;

            try
            {
                IEnumerable<DatosTablaDto> datos = (IEnumerable<DatosTablaDto>)RecuperarDatosParaTabla();
                IEnumerable<DatoTablaDto> datosGrafica = datos.First().DatoTablaDtosList;

                grafica = (NetOffice.PowerPointApi.Chart)diapositiva.Shapes[NumElemGrafica[1]].Chart;
                graficaInfo = grafica.ChartData;

                cuaderno = (NetOffice.ExcelApi._Workbook)graficaInfo.Workbook;
                hoja = (NetOffice.ExcelApi._Worksheet)cuaderno.Worksheets[9];                

                IEnumerable<String> anios = datosGrafica.Where(x => x.Concepto.Equals(ConstantesBl.TOTAL_GENERAL) && !(x.Titulo.Equals(ConstantesBl.TOTAL))).Select(x => x.Titulo).Distinct();
                int numRenglon = 3;
                int cAnios = 0;

                if (datosGrafica != null && datosGrafica.Count() > 0)
                {
                    cAnios = 0;
                    foreach (String anio in anios)
                    {
                        if (cAnios >= 2 || anios.Count() <= 3)
                        {
                            ((NetOffice.ExcelApi.Range)hoja.Cells.get_Range(string.Format("B{0}", numRenglon), Type.Missing)).FormulaR1C1 = TipoDatoBase.FormatearDato(TipoDatoAtrib.TEXTO, anio);

                            ((NetOffice.ExcelApi.Range)hoja.Cells.get_Range(string.Format("C{0}", numRenglon), Type.Missing)).FormulaR1C1 = TipoDatoBase.FormatearDato(TipoDatoAtrib.ENTERO, datosGrafica.Where(x => x.Titulo.Equals(anio) && x.Concepto.Equals(ConstantesBl.TOTAL_GENERAL)).Select(x => x.Importe).Sum());
                            numRenglon++;
                        }
                        cAnios++;
                    }
                }

                grafica.ApplyDataLabels();
                grafica.Refresh();


            }
            catch (Exception ex)
            {
                Excepcion.InsertarException(ex, CodUsuario);
            }
            finally
            {
                if (cuaderno != null)
                {
                    cuaderno.Save();
                    cuaderno.Close(true);
                    grafica.Application.Quit();
                    grafica.Dispose();
                }
            }

            return diapositiva;
        }

        #endregion Metodos Privados
    }
}
