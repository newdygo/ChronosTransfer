﻿using System;
using System.Web;
using System.Data;
using System.Linq;
using System.Data.OleDb;
using System.Globalization;
using System.Web.UI.WebControls;
using System.Collections.Generic;
using System.Reflection;
using Excels = Microsoft.Office.Interop.Excel;
using System.IO;
using System.Diagnostics;

namespace ChronosTransfer.CLTransfer
{
    public class Passageiro : Chronos
    {
        #region Propriedades

        #region Básico

        public String Nome { get; set; }
        public String DocumentoIdentificacao { get; set; }

        #endregion

        #region Chegada

        public DateTime DataC { get; set; }
        public String CidadeOrigemC { get; set; }
        public String CidadeDestinoC { get; set; }
        public String CompanhiaAereaC { get; set; }
        public String NumeroVooC { get; set; }
        public DateTime HorarioSaidaC { get; set; }
        public DateTime HorarioChegadaC { get; set; }

        #endregion

        #region Partida

        public DateTime DataP { get; set; }
        public String CidadeOrigemP { get; set; }
        public String CidadeDestinoP { get; set; }
        public String CompanhiaAereaP { get; set; }
        public String NumeroVooP { get; set; }
        public DateTime HorarioSaidaP { get; set; }
        public DateTime HorarioChegadaP { get; set; }

        #endregion

        #endregion

        #region Propriedades Privadas

        DataTable _Table;
        Passageiro _Passageiro;
        List<Passageiro> _Passageiros;
        List<Passageiro> _PassageirosOrdenadorNumeroVoo;
        List<VooChegada> _TransporteVooChegada;

        Char[] _Char = Environment.NewLine.ToCharArray();

        #endregion

        #region Métodos

        /// <summary>
        /// Rotina utilizada para processar o arquivo, criando uma lista ordenada e percorrendo agrupando os valores
        /// </summary>
        /// <param name="_DataSource"></param>
        /// <param name="_Sheet"></param>
        /// <returns></returns>
        public List<VooChegada> ProcessarArquivo(String _DataSource, String _Sheet)
        {
            _Passageiros = new List<Passageiro>();

            _Table = new Excel().RetornarTudoSheet(_Sheet);

            DataTable _ttt = RetornarColumnsSchema(OleDbSchemaGuid.Columns, _Sheet);

            foreach (DataRow _Row in _Table.Rows)
            {
                try
                {
                    _Passageiro = new Passageiro();

                    IFormatProvider _Culture = new System.Globalization.CultureInfo("pt-BR", true);

                    _Passageiro.Nome = _Row[0].ToString();
                    _Passageiro.DocumentoIdentificacao = _Row[1].ToString();

                    _Passageiro.DataC = DateTime.Parse(_Row[2].ToString().Split(_Char).Last().ToString(), _Culture, DateTimeStyles.None);
                    _Passageiro.CidadeOrigemC = _Row[3].ToString().Split(_Char).Last().ToString();
                    _Passageiro.CidadeDestinoC = _Row[4].ToString().Split(_Char).Last().ToString();
                    _Passageiro.CompanhiaAereaC = _Row[5].ToString().Split(_Char).Last().ToString();
                    _Passageiro.NumeroVooC = _Row[6].ToString().Split(_Char).Last().ToString();
                    _Passageiro.HorarioSaidaC = DateTime.Parse(_Row[7].ToString().Split(_Char).Last().ToString(), _Culture, DateTimeStyles.None);
                    _Passageiro.HorarioChegadaC = DateTime.Parse(_Row[8].ToString().Split(_Char).Last().ToString(), _Culture, DateTimeStyles.None);

                    _Passageiro.DataP = DateTime.Parse(_Row[10].ToString().Split(_Char).Last().ToString(), _Culture, DateTimeStyles.None);
                    _Passageiro.CidadeOrigemP = _Row[11].ToString().Split(_Char).Last().ToString();
                    _Passageiro.CidadeDestinoP = _Row[12].ToString().Split(_Char).Last().ToString();
                    _Passageiro.CompanhiaAereaP = _Row[13].ToString().Split(_Char).Last().ToString();
                    _Passageiro.NumeroVooP = _Row[14].ToString().Split(_Char).Last().ToString();
                    _Passageiro.HorarioSaidaP = DateTime.Parse(_Row[15].ToString().Split(_Char).Last().ToString(), _Culture, DateTimeStyles.None);
                    _Passageiro.HorarioChegadaP = DateTime.Parse(_Row[16].ToString().Split(_Char).Last().ToString(), _Culture, DateTimeStyles.None);

                    _Passageiros.Add(_Passageiro);
                }
                catch
                {
                    continue;
                }
            }

            _PassageirosOrdenadorNumeroVoo = _Passageiros.OrderBy(x => x.NumeroVooC).ToList();

            _TransporteVooChegada = new List<VooChegada>();

            foreach (Passageiro _Passageiro in _PassageirosOrdenadorNumeroVoo)
            {
                if (_TransporteVooChegada.Count == 0)
                {
                    VooChegada _voo = new VooChegada();

                    _voo.NumeroVoo = _Passageiro.NumeroVooC;
                    _voo.Data = _Passageiro.DataC;
                    _voo.HorarioSaida = _Passageiro.HorarioSaidaC;
                    _voo.HorarioChegada = _Passageiro.HorarioChegadaC;
                    _voo.QuantidadePassageiro += 1;
                    _voo.TipoVeiculo = RetornaCarro(_voo.QuantidadePassageiro);

                    _TransporteVooChegada.Add(_voo);
                }
                else
                {
                    if (_TransporteVooChegada.Last().NumeroVoo == _Passageiro.NumeroVooC)
                    {
                        _TransporteVooChegada.Last().QuantidadePassageiro += 1;
                        _TransporteVooChegada.Last().TipoVeiculo = RetornaCarro(_TransporteVooChegada.Last().QuantidadePassageiro);
                    }
                    else
                    {
                        VooChegada _voo = new VooChegada();

                        _voo.NumeroVoo = _Passageiro.NumeroVooC;
                        _voo.Data = _Passageiro.DataC;
                        _voo.HorarioSaida = _Passageiro.HorarioSaidaC;
                        _voo.HorarioChegada = _Passageiro.HorarioChegadaC;
                        _voo.QuantidadePassageiro += 1;
                        _voo.TipoVeiculo = RetornaCarro(_voo.QuantidadePassageiro);

                        _TransporteVooChegada.Add(_voo);
                    }
                }
            }

            return _TransporteVooChegada;

            //DataTable _Table2 = new DataTable();

            //try
            //{

            //    Passageiro Pax = new Passageiro();

            //    List<TransporteVooChegada> _Tra = new List<TransporteVooChegada>();
            //    List<TransporteVooChegada> _Tra3 = new List<TransporteVooChegada>();


            //    List<TransporteVooChegada> _Tra2 = _Tra.OrderBy(x => x.HorarioChegada).ToList();

            //    foreach (TransporteVooChegada passa in _Tra2)
            //    {
            //        List<TransporteVooChegada> _ooo = new List<TransporteVooChegada>();

            //        _ooo = _Tra2.Where(x => x.HorarioChegada <= passa.HorarioChegada.AddMinutes(90) && x.HorarioChegada > passa.HorarioChegada.AddMinutes(-90)).ToList();

            //        TransporteVooChegada ppp = new TransporteVooChegada();
            //        foreach (var ooo in _ooo)
            //        {
            //            if (ppp.NumeroVoo == String.Empty)
            //            {
            //                ppp.NumeroVoo = String.Format("Voo: {0}", ooo.NumeroVoo);
            //            }
            //            else
            //            {
            //                ppp.NumeroVoo = String.Format("{0} - Voo: {1}", ppp.NumeroVoo, ooo.NumeroVoo);
            //            }

            //            if (ppp.HorarioChegada < ooo.HorarioChegada)
            //                ppp.HorarioChegada = ooo.HorarioChegada;

            //            if (ppp.HorarioSaida < ooo.HorarioSaida)
            //                ppp.HorarioSaida = ooo.HorarioSaida;

            //            ppp.Quantidade += ooo.Quantidade;
            //            ppp.TipoVeiculo = RetornaCarro(ppp.Quantidade);
            //        }

            //        _Tra3.Add(ppp);
            //    }

            //    _GridView.DataSource = _Tra2;

            //    _GridView.DataBind();

            //}
            //catch (Exception ex)
            //{                
            //    throw ex;
            //}
        }        

        /// <summary>
        /// Rotina utilizada para identificar qual o veiculo de acordo com a quantidade de pessoas
        /// </summary>
        /// <param name="Quantidade"></param>
        /// <returns></returns>
        public String RetornaCarro(int Quantidade)
        {
            if (Quantidade == 1)
            {
                return "Moto";
            }
            else if (Quantidade == 2)
            {
                return "Triciculo";
            }
            else if (Quantidade == 3)
            {
                return "Carro";
            }
            else if (Quantidade > 3 && Quantidade <= 8)
            {
                return "Mini Van";
            }
            else if (Quantidade > 8 && Quantidade <= 11)
            {
                return "Mini ônibus";
            }
            else
            {
                return "Ônibus";
            }
        }

        public void Conectar2()
        {
            try
            {
                Process myProcess = new Process();

                myProcess.StartInfo.FileName = @"C:\WINDOWS\system32\regsvr32.exe";
                myProcess.StartInfo.Arguments = "COM.DLL";
                myProcess.StartInfo.CreateNoWindow = true;
                myProcess.Start();

                Excels.Application myApp;
                Excels.Workbook myWorkBk;
                object missingValue = System.Reflection.Missing.Value;
                myApp = new Excels.Application();
                myWorkBk = myApp.Workbooks.Add(missingValue);

                Excels.Worksheet myWorkSht;
                myWorkSht = (Excels.Worksheet)myWorkBk.Worksheets.get_Item(1);

                myWorkSht.Name = "Students-Information";

                //myWorkSht.Cells[1, 1] = "Let us introduce our Great Students";

                //myWorkSht.Cells[4, 6] = "=AVERAGE(C4:E4)";

                //myWorkSht.get_Range("B3", "E3").Font.Size = 14;
                //myWorkSht.get_Range("B3", "E3").Font.Bold = true;
                //myWorkSht.get_Range("B3", "E3").HorizontalAlignment = Excels.XlVAlign.xlVAlignCenter;
                //myWorkSht.get_Range("B3", "E3").Interior.Color = System.Drawing.Color.LightSkyBlue;

                //myWorkSht.get_Range("B3", "E3").Borders.LineStyle = 1;
                //myWorkSht.get_Range("B4", "B6").Borders[Excels.XlBordersIndex.xlEdgeLeft].Weight = Excels.XlBorderWeight.xlThin;
                //myWorkSht.get_Range("E4", "E6").Borders[Excels.XlBordersIndex.xlEdgeRight].Weight = Excels.XlBorderWeight.xlThin;
                //myWorkSht.get_Range("B4", "E4").Borders[Excels.XlBordersIndex.xlEdgeBottom].Weight = Excels.XlBorderWeight.xlThin;
                //myWorkSht.get_Range("B3", "C3").Borders[Excels.XlBordersIndex.xlEdgeTop].Weight = Excels.XlBorderWeight.xlThin;

                //myWorkSht.Columns[4].ColumnWidth = 13.5;

                myWorkBk.SaveAs(Path.GetTempPath() + "Diego.xls");   
            }
            catch (Exception ex)
            {                
                throw ex;
            }           
        }

        #endregion
    }
}