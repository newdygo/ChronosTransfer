using System;
using System.IO;
using System.Web;
using System.Data;
using System.Linq;
using OfficeOpenXml;
using System.Data.OleDb;
using System.Collections.Generic;
using ChronosTransfer.CLChronos.CLTransfer;
using ChronosTransfer.CLChronos.CLChronosRaiz.CLChronos;
using OfficeOpenXml.Style;
using ChronosTransfer.CLChronos.CLPassageiro;

namespace ChronosTransfer.CLChronos.CLExcel
{
    public class Excel : Chronos
    {
        #region Construtores
        
        public Excel(String _Source) : base(_Source) { }
        
        #endregion

        #region Métodos

        #region Básicos
        
        /// <summary>
        /// Rotina utilizada para retornar todos os dados de uma sheet
        /// </summary>
        /// <param name="_Sheet">Sheet da planilha Excel.</param>
        /// <returns>Retorna dados de uma sheet Excel.</returns>
        public DataTable RetornarTudoSheet(String _Sheet)
        {
            DataTable tbl = new DataTable();

            using (ConectarEx())
            {
                ExcelWorkbook _WorkBook = ConexaoEx.Workbook;

                var ws = ConexaoEx.Workbook.Worksheets[_Sheet];                
                bool hasHeader = true; // adjust it accordingly( i've mentioned that this is a simple approach)
                foreach (var firstRowCell in ws.Cells[1, 1, 1, ws.Dimension.End.Column])
                {
                    if (tbl.Columns.Contains(firstRowCell.Text))
                    {
                        tbl.Columns.Add(String.Format("{0}_2", firstRowCell.Text));
                    }
                    else
                    {
                        tbl.Columns.Add(firstRowCell.Text);
                    }                    
                }

                var startRow = hasHeader ? 2 : 1;
                for (var rowNum = startRow; rowNum <= ws.Dimension.End.Row; rowNum++)
                {
                    var wsRow = ws.Cells[rowNum, 1, rowNum, ws.Dimension.End.Column];
                    var row = tbl.NewRow();
                    foreach (var cell in wsRow)
                    {
                        row[cell.Start.Column - 1] = cell.Text;
                    }
                    tbl.Rows.Add(row);
                }
            }

            return tbl;
        }

        #endregion

        #region Úteis
        
        /// <summary>
        /// Rotina utilizada para retornar as Sheets da planilha.
        /// </summary>
        /// <param name="_DataSource">Caminho do arquivo de conexão com o Excel.</param>
        /// <returns>Retorn um datatable com os nomes das sheets da planilha Excel.</returns>
        public List<String> RetornarSchema()
        {
            using (ConectarEx())
            {
                return (from Sheet in ConexaoEx.Workbook.Worksheets where Sheet.Hidden == eWorkSheetHidden.Visible select Sheet.Name).ToList();
            }
        }

        #endregion

        #endregion

        /// <summary>
        /// Rotina utilizada para criar uma Sheet num excel.
        /// </summary>
        /// <param name="_Sheet"></param>
        /// <param name="_Passageiros"></param>
        public void CreateSheet(String _Sheet)
        {
            using (ConectarEx())
            {
                ExcelWorkbook _WorkBook = ConexaoEx.Workbook;

                if (_WorkBook.Worksheets.Where(x => String.Format("TransferChegada_{0}", _Sheet) == x.Name).Count() > 0)
                {
                    _WorkBook.Worksheets.Delete(_WorkBook.Worksheets.Where(x => String.Format("TransferChegada_{0}", _Sheet) == x.Name).First());
                }

                if (_WorkBook.Worksheets.Where(x => String.Format("TransferRetorno_{0}", _Sheet) == x.Name).Count() > 0)
                {
                    _WorkBook.Worksheets.Delete(_WorkBook.Worksheets.Where(x => String.Format("TransferRetorno_{0}", _Sheet) == x.Name).First());
                }

                _WorkBook.Worksheets.Add(String.Format("TransferChegada_{0}", _Sheet));
                ExcelWorksheet _WorkSheet = _WorkBook.Worksheets.Last();

                _WorkSheet.Cells["A1"].Value = "Transfer";
                _WorkSheet.Cells["B1"].Value = "Nome";
                _WorkSheet.Cells["C1"].Value = "Documento";
                _WorkSheet.Cells["D1"].Value = "Data";
                _WorkSheet.Cells["E1"].Value = "Cidade Origem";
                _WorkSheet.Cells["F1"].Value = "Cidade Destino";
                _WorkSheet.Cells["G1"].Value = "Número Voô";
                _WorkSheet.Cells["H1"].Value = "Horário Saída";
                _WorkSheet.Cells["I1"].Value = "Horário Chegada";
                _WorkSheet.Cells["J1"].Value = "Veículo";
                _WorkSheet.Cells["K1"].Value = "Valor";

                _WorkSheet.Cells["A1:K1"].Style.Font.Bold = true;
                _WorkSheet.Cells["A1:K1"].Style.Font.Size = 12;
                _WorkSheet.Cells["A1:K1"].AutoFitColumns();

                _WorkBook.Worksheets.Add(String.Format("TransferRetorno_{0}", _Sheet));
                _WorkSheet = _WorkBook.Worksheets.Last();

                _WorkSheet.Cells["A1"].Value = "Transfer";
                _WorkSheet.Cells["B1"].Value = "Nome";
                _WorkSheet.Cells["C1"].Value = "Documento";
                _WorkSheet.Cells["D1"].Value = "Data";
                _WorkSheet.Cells["E1"].Value = "Cidade Origem";
                _WorkSheet.Cells["F1"].Value = "Cidade Destino";
                _WorkSheet.Cells["G1"].Value = "Número Voô";
                _WorkSheet.Cells["H1"].Value = "Horário Saída";
                _WorkSheet.Cells["I1"].Value = "Horário Chegada";
                _WorkSheet.Cells["J1"].Value = "Veículo";
                _WorkSheet.Cells["K1"].Value = "Valor";

                _WorkSheet.Cells["A1:K1"].Style.Font.Bold = true;
                _WorkSheet.Cells["A1:K1"].Style.Font.Size = 12;
                _WorkSheet.Cells["A1:K1"].AutoFitColumns();

                if (ConexaoEx.Stream.CanWrite)
                {
                    ConexaoEx.Save();
                }                
            }
        }

        /// <summary>
        /// Rotina utilizada para inserir voo de chegada a sheet do Excel.
        /// </summary>
        /// <param name="_Sheet"></param>
        /// <param name="_Transfer"></param>
        public void InsertPassageiroSheetChegada(String _Sheet, Transfer _Transfer)
        {
            using (ConectarEx())
            {
                ExcelWorkbook _WorkBook = ConexaoEx.Workbook;

                ExcelWorksheet _WorkSheet = _WorkBook.Worksheets.Where(x => String.Format("TransferChegada_{0}", _Sheet) == x.Name).First();

                int _Start = _WorkSheet.Dimension.End.Row;
                int _StartE = _WorkSheet.Dimension.End.Row;

                foreach (Passageiro _Passageiro in _Transfer.Passageiros)
                {
                    _Start = _Start + 1;

                    _WorkSheet.InsertRow(_Start, 1);

                    if (_Transfer.Passageiros.IndexOf(_Passageiro) == _Transfer.Passageiros.Count -1)
                    {
                        _StartE = _StartE + 1;

                        String _MergeA = String.Format("A{0}:A{1}", _StartE, _StartE + _Transfer.Passageiros.Count - 1);
                        String _MergeJ = String.Format("J{0}:J{1}", _StartE, _StartE + _Transfer.Passageiros.Count - 1);
                        String _MergeK = String.Format("K{0}:K{1}", _StartE, _StartE + _Transfer.Passageiros.Count - 1);

                        _WorkSheet.Cells[_MergeA].Merge = true;
                        _WorkSheet.Cells[_MergeA].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        _WorkSheet.Cells[_MergeA].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                        _WorkSheet.Cells[String.Format("A{0}", _StartE)].Value = _Transfer.Nome;

                        _WorkSheet.Cells[_MergeJ].Merge = true;
                        _WorkSheet.Cells[_MergeJ].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        _WorkSheet.Cells[_MergeJ].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                        _WorkSheet.Cells[String.Format("J{0}", _StartE)].Value = _Transfer.Veiculo.Nome;

                        _WorkSheet.Cells[_MergeK].Merge = true;
                        _WorkSheet.Cells[_MergeK].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        _WorkSheet.Cells[_MergeK].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                        _WorkSheet.Cells[String.Format("K{0}", _StartE)].Value = _Transfer.Veiculo.ValorViagem;
                        _WorkSheet.Cells[String.Format("K{0}", _StartE)].Style.Numberformat.Format = @"R$ #,##0.00";
                    }
                    
                    _WorkSheet.Cells[String.Format("B{0}", _Start)].Value = _Passageiro.Nome;
                    _WorkSheet.Cells[String.Format("C{0}", _Start)].Value = _Passageiro.Documento;
                    _WorkSheet.Cells[String.Format("D{0}", _Start)].Value = _Passageiro.VooChegada.NumeroVoo;
                    _WorkSheet.Cells[String.Format("E{0}", _Start)].Value = _Passageiro.VooChegada.CidadeOrigem;
                    _WorkSheet.Cells[String.Format("F{0}", _Start)].Value = _Passageiro.VooChegada.CidadeDestino;
                    _WorkSheet.Cells[String.Format("H{0}", _Start)].Value = _Passageiro.VooChegada.HorarioSaida.ToShortTimeString();
                    _WorkSheet.Cells[String.Format("I{0}", _Start)].Value = _Passageiro.VooChegada.HorarioChegada.ToShortTimeString();
                }

                if (ConexaoEx.Stream.CanWrite)
                {
                    ConexaoEx.Save();
                }  
            }
        }

        /// <summary>
        /// Rotina utilizada para inserir voo de chegada a sheet do Excel.
        /// </summary>
        /// <param name="_Sheet"></param>
        /// <param name="_Transfer"></param>
        public void InsertPassageiroSheetRetorno(String _Sheet, Transfer _Transfer)
        {
            using (ConectarEx())
            {
                ExcelWorkbook _WorkBook = ConexaoEx.Workbook;

                ExcelWorksheet _WorkSheet = _WorkBook.Worksheets.Where(x => String.Format("TransferRetorno_{0}", _Sheet) == x.Name).First();

                int _Start = _WorkSheet.Dimension.End.Row;
                int _StartE = _WorkSheet.Dimension.End.Row;

                foreach (Passageiro _Passageiro in _Transfer.Passageiros)
                {
                    _Start = _Start + 1;

                    _WorkSheet.InsertRow(_Start, 1);

                    if (_Transfer.Passageiros.IndexOf(_Passageiro) == _Transfer.Passageiros.Count -1)
                    {
                        _StartE = _StartE + 1;

                        String _MergeA = String.Format("A{0}:A{1}", _StartE, _StartE + _Transfer.Passageiros.Count - 1);
                        String _MergeJ = String.Format("J{0}:J{1}", _StartE, _StartE + _Transfer.Passageiros.Count - 1);
                        String _MergeK = String.Format("K{0}:K{1}", _StartE, _StartE + _Transfer.Passageiros.Count - 1);

                        _WorkSheet.Cells[_MergeA].Merge = true;
                        _WorkSheet.Cells[_MergeA].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        _WorkSheet.Cells[_MergeA].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                        _WorkSheet.Cells[String.Format("A{0}", _StartE)].Value = _Transfer.Nome;

                        _WorkSheet.Cells[_MergeJ].Merge = true;
                        _WorkSheet.Cells[_MergeJ].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        _WorkSheet.Cells[_MergeJ].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                        _WorkSheet.Cells[String.Format("J{0}", _StartE)].Value = _Transfer.Veiculo.Nome;

                        _WorkSheet.Cells[_MergeK].Merge = true;
                        _WorkSheet.Cells[_MergeK].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        _WorkSheet.Cells[_MergeK].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                        _WorkSheet.Cells[String.Format("K{0}", _StartE)].Value = _Transfer.Veiculo.ValorViagem;
                        _WorkSheet.Cells[String.Format("K{0}", _StartE)].Style.Numberformat.Format = @"R$ #,##0.00";
                    }

                    _WorkSheet.Cells[String.Format("B{0}", _Start)].Value = _Passageiro.Nome;
                    _WorkSheet.Cells[String.Format("C{0}", _Start)].Value = _Passageiro.Documento;
                    _WorkSheet.Cells[String.Format("D{0}", _Start)].Value = _Passageiro.VooRetorno.NumeroVoo;
                    _WorkSheet.Cells[String.Format("E{0}", _Start)].Value = _Passageiro.VooRetorno.CidadeOrigem;
                    _WorkSheet.Cells[String.Format("F{0}", _Start)].Value = _Passageiro.VooRetorno.CidadeDestino;
                    _WorkSheet.Cells[String.Format("H{0}", _Start)].Value = _Passageiro.VooRetorno.HorarioSaida.ToShortTimeString();
                    _WorkSheet.Cells[String.Format("I{0}", _Start)].Value = _Passageiro.VooRetorno.HorarioChegada.ToShortTimeString();
                }

                if (ConexaoEx.Stream.CanWrite)
                {
                    ConexaoEx.Save();
                }  
            }
        }
    }
}