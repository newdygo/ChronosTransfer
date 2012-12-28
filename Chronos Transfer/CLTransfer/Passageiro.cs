using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Linq;
using System.Web;
using System.Web.UI.WebControls;

namespace Chronos_Transfer.CLTransfer
{
    public class Passageiro
    {
         #region Variáveis

        private String _Nome;
        private String _DocumentoIdentificacao;

        #region Chegada

        private DateTime _DataC;
        private String _CidadeOrigemC;
        private String _CidadeDestinoC;
        private String _CompanhiaAereaC;
        private String _NumeroVooC;
        private DateTime _HorarioSaidaC;
        private DateTime _HorarioChegadaC;

        #endregion

        #region Partida

        private DateTime _DataP;
        private String _CidadeOrigemP;
        private String _CidadeDestinoP;
        private String _CompanhiaAereaP;
        private String _NumeroVooP;
        private DateTime _HorarioSaidaP;
        private DateTime _HorarioChegadaP;

        #endregion

        #endregion

        #region Propriedades

        public String Nome { get; set; }
        public String DocumentoIdentificacao { get; set; }

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

        public void AbrirArquivo(String _DataSource, GridView _GridView)
        {
            List<Passageiro> _Passageiros = new List<Passageiro>();
            DataTable _Table = new DataTable();

            try
            {
                String Cone = String.Format("Provider=Microsoft.Jet.OleDb.4.0; data source= {0}; Extended Properties=\"Excel 8.0; HDR=Yes\";", _DataSource);

                using (OleDbConnection connection = new OleDbConnection(Cone))
                {
                    // The insertSQL string contains a SQL statement that
                    // inserts a new row in the source table.
                    OleDbCommand command = new OleDbCommand("select * from [Plan1$] where Nome <> ''");

                    command.Connection = connection;

                    try
                    {
                        connection.Open();                       

                        _Table.Load(command.ExecuteReader());
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex.Message);
                    }
                }
                int valor = 0;

                foreach (DataRow Linha in _Table.Rows)
                {
                    try
                    {
                        Passageiro Pax = new Passageiro();

                        Pax.Nome = Linha[0].ToString();
                        Pax.DocumentoIdentificacao = Linha[1].ToString();

                        Pax.DataC = Convert.ToDateTime(Linha[2].ToString().Split('\n').Last().ToString());
                        Pax.CidadeOrigemC = Linha[3].ToString().Split('\n').Last().ToString();
                        Pax.CidadeDestinoC = Linha[4].ToString().Split('\n').Last().ToString();
                        Pax.CompanhiaAereaC = Linha[5].ToString().Split('\n').Last().ToString();
                        Pax.NumeroVooC = Linha[6].ToString().Split('\n').Last().ToString();
                        Pax.HorarioSaidaC = Convert.ToDateTime(Linha[7].ToString().Split('\n').First().ToString());
                        Pax.HorarioChegadaC = Convert.ToDateTime(Linha[8].ToString().Split('\n').Last().ToString());

                        Pax.DataP = Convert.ToDateTime(Linha[10].ToString().Split('\n').Last().ToString());
                        Pax.CidadeOrigemP = Linha[11].ToString().Split('\n').Last().ToString();
                        Pax.CidadeDestinoP = Linha[12].ToString().Split('\n').Last().ToString();
                        Pax.CompanhiaAereaP = Linha[13].ToString().Split('\n').Last().ToString();
                        Pax.NumeroVooP = Linha[14].ToString().Split('\n').Last().ToString();
                        Pax.HorarioSaidaP = Convert.ToDateTime(Linha[15].ToString().Split('\n').First().ToString());
                        Pax.HorarioChegadaP = Convert.ToDateTime(Linha[16].ToString().Split('\n').Last().ToString());

                        _Passageiros.Add(Pax);

                        valor = _Table.Rows.IndexOf(Linha);
                    }
                    catch (Exception ex)
                    {   
                        throw;
                    }                    
                }

                List<Passageiro> _Pasa = _Passageiros.OrderBy(x => x.NumeroVooC).ToList();
                List<TransporteVooChegada> _Tra = new List<TransporteVooChegada>();
                List<TransporteVooChegada> _Tra3 = new List<TransporteVooChegada>();

                foreach (Passageiro passa in _Pasa)
                {
                    if (_Tra.Count == 0)
                    {
                        TransporteVooChegada _voo = new TransporteVooChegada();

                        _voo.NumeroVoo = passa.NumeroVooC;
                        _voo.HorarioSaida = passa.HorarioSaidaC;
                        _voo.HorarioChegada = passa.HorarioChegadaC;
                        _voo.Quantidade += 1;
                        _voo.TipoVeiculo = RetornaCarro(_voo.Quantidade);

                        _Tra.Add(_voo);
                    }
                    else
                    {
                        if (_Tra.Last().NumeroVoo == passa.NumeroVooC)
                        {
                            _Tra.Last().Quantidade += 1;
                            _Tra.Last().TipoVeiculo = RetornaCarro(_Tra.Last().Quantidade);
                        }
                        else
                        {
                            TransporteVooChegada _voo = new TransporteVooChegada();

                            _voo.NumeroVoo = passa.NumeroVooC;
                            _voo.HorarioSaida = passa.HorarioSaidaC;
                            _voo.HorarioChegada = passa.HorarioChegadaC;
                            _voo.Quantidade += 1;
                            _voo.TipoVeiculo = RetornaCarro(_voo.Quantidade);

                            _Tra.Add(_voo);
                        }
                    }
                }

                List<TransporteVooChegada> _Tra2 = _Tra.OrderBy(x => x.HorarioChegada).ToList();

                foreach (TransporteVooChegada passa in _Tra2)
                {
                    List<TransporteVooChegada> _ooo = new List<TransporteVooChegada>();

                    _ooo = _Tra2.Where(x => x.HorarioChegada <= passa.HorarioChegada.AddMinutes(90) && x.HorarioChegada > passa.HorarioChegada.AddMinutes(-90)).ToList();

                    TransporteVooChegada ppp = new TransporteVooChegada();
                    foreach (var ooo in _ooo)
                    {
                        if (ppp.NumeroVoo == String.Empty)
                        {
                            ppp.NumeroVoo = String.Format("Voo: {0}", ooo.NumeroVoo);
                        }
                        else
                        {
                            ppp.NumeroVoo = String.Format("{0} - Voo: {1}", ppp.NumeroVoo, ooo.NumeroVoo);
                        }

                        if (ppp.HorarioChegada < ooo.HorarioChegada)
                            ppp.HorarioChegada = ooo.HorarioChegada;

                        if (ppp.HorarioSaida < ooo.HorarioSaida)
                            ppp.HorarioSaida = ooo.HorarioSaida;

                        ppp.Quantidade += ooo.Quantidade;
                        ppp.TipoVeiculo = RetornaCarro(ppp.Quantidade);
                    }

                    _Tra3.Add(ppp);
                }
                
                //BindingSource _BBB = new BindingSource();

                _GridView.DataSource = _Tra2;

                _GridView.DataBind();

                //_1.DataSource = _BBB;

                //BindingSource _BBB2 = new BindingSource();

                //_BBB2.DataSource = _Tra3;

                //_2.DataSource = _BBB2;
       
            }
            catch (Exception ex)
            {                
                throw ex;
            }
        }

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
    }
}