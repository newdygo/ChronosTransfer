using System;
using System.Web;
using System.Data;
using System.Linq;
using System.Data.OleDb;
using System.Globalization;
using System.Web.UI.WebControls;
using System.Collections.Generic;
using ChronosTransfer.CLChronos.CLBasico;
using ChronosTransfer.CLChronos.CLPassageiro;
using ChronosTransfer.CLChronos.CLChronosRaiz.CLChronos;
using ChronosTransfer.CLChronos.CLExcel;
using ChronosTransfer.CLChronos.CLVoo;

namespace ChronosTransfer.CLChronos.CLTransfer
{
    public class Transfer : Chronos
    {
        #region Contrutores

        public Transfer() : base() {}
        
        public Transfer(String _Source) : base(_Source) {}

        #endregion

        #region Propriedades

        private String Sheet { get; set; }

        public List<Passageiro> Passageiros { get; set; }
        public Veiculo Veiculo { get; set; }    
        
        #endregion

        #region Métodos

        public List<Transfer> GerarTransfers(String _Sheet)
        {
            Passageiros = new List<Passageiro>();

            List<Veiculo> _Veiculos = new List<Veiculo>();

            _Veiculos.Add(new Veiculo() { Nome = "Moto", Capacidade = 1, ValorViagem = Convert.ToDecimal("350.00") });
            _Veiculos.Add(new Veiculo() { Nome = "Carro", Capacidade = 4, ValorViagem = Convert.ToDecimal("350.00") });
            _Veiculos.Add(new Veiculo() { Nome = "Mini Van", Capacidade = 8, ValorViagem = Convert.ToDecimal("700.00") });
            _Veiculos.Add(new Veiculo() { Nome = "Micro Ônibus", Capacidade = 22, ValorViagem = Convert.ToDecimal("1000.00") });
            _Veiculos.Add(new Veiculo() { Nome = "Ônibus", Capacidade = 38, ValorViagem = Convert.ToDecimal("1500.00") });

            DataTable _Table = new Excel().RetornarTudoSheet(_Sheet);
            
            Passageiro _Passageiro = new Passageiro();

            Voo _VooChegada = new Voo();
            Voo _VooRetorno = new Voo();

            Char[] _Char = Environment.NewLine.ToCharArray();
         
            foreach (DataRow _Row in _Table.Rows)
            {
                try
                {
                    IFormatProvider _Culture = new System.Globalization.CultureInfo("pt-BR", true);

                    _Passageiro.Nome = _Row[0].ToString();
                    _Passageiro.Documento = _Row[1].ToString();
                    
                    _VooChegada.Data = DateTime.Parse(_Row[2].ToString().Split(_Char).Last().ToString(), _Culture, DateTimeStyles.None);
                    _VooChegada.CidadeOrigem = _Row[3].ToString().Split(_Char).Last().ToString();
                    _VooChegada.CidadeDestino = _Row[4].ToString().Split(_Char).Last().ToString();
                    _VooChegada.CompanhiaAerea = _Row[5].ToString().Split(_Char).Last().ToString();
                    _VooChegada.NumeroVoo = _Row[6].ToString().Split(_Char).Last().ToString();
                    _VooChegada.HorarioSaida = DateTime.Parse(_Row[7].ToString().Split(_Char).Last().ToString(), _Culture, DateTimeStyles.None);
                    _VooChegada.HorarioChegada = DateTime.Parse(_Row[8].ToString().Split(_Char).Last().ToString(), _Culture, DateTimeStyles.None);

                    _VooRetorno.Data = DateTime.Parse(_Row[10].ToString().Split(_Char).Last().ToString(), _Culture, DateTimeStyles.None);
                    _VooRetorno.CidadeOrigem = _Row[11].ToString().Split(_Char).Last().ToString();
                    _VooRetorno.CidadeDestino = _Row[12].ToString().Split(_Char).Last().ToString();
                    _VooRetorno.CompanhiaAerea = _Row[13].ToString().Split(_Char).Last().ToString();
                    _VooRetorno.NumeroVoo = _Row[14].ToString().Split(_Char).Last().ToString();
                    _VooRetorno.HorarioSaida = DateTime.Parse(_Row[15].ToString().Split(_Char).Last().ToString(), _Culture, DateTimeStyles.None);
                    _VooRetorno.HorarioChegada = DateTime.Parse(_Row[16].ToString().Split(_Char).Last().ToString(), _Culture, DateTimeStyles.None);

                    _Passageiro.VooChegada = _VooChegada;
                    _Passageiro.VooRetorno = _VooRetorno;

                    Passageiros.Add(_Passageiro);

                    _Passageiro = new Passageiro();
                    _VooChegada = new Voo();
                    _VooRetorno = new Voo();
                }
                catch
                {
                    //Gerar Log do ocorrido e em qual linha e continuar, no final verificar o logo e exibir para o usuário corrigir as pendências.
                    //Não exibir o resultado caso o log seja maior que 0. exibir o log.
                    continue;
                }
            }

            List<Passageiro> _PassageirosOrdenados = Passageiros.OrderBy(x => x.VooChegada.NumeroVoo).ToList();
            
            List<Transfer> _Transfers = new List<Transfer>();
                        
            foreach (IGrouping<String, Passageiro> _PassageiroOrdenado in _PassageirosOrdenados.GroupBy(x => x.VooChegada.NumeroVoo))
            {
                Transfer _Transfer = new Transfer();

                if (_PassageiroOrdenado.Count() > _Veiculos.Max(x => x.Capacidade))
                {
                    List<Passageiro> _PassageirosTemp = _PassageiroOrdenado.ToList();
                    
                    for (int i = _PassageirosTemp.Count / _Veiculos.Max(x => x.Capacidade); i >= 0; i--)
                    {   
                        _Transfer.Passageiros = _PassageirosTemp.Take(38).ToList();
                        _Transfer.Veiculo = new Veiculo(_Transfer.Passageiros.Count);

                        _Transfers.Add(_Transfer);

                        _PassageirosTemp.RemoveRange(0, _Transfer.Passageiros.Count);

                        _Transfer = new Transfer();
                    }                    
                }
                else
                {
                    _Transfer.Passageiros = _PassageiroOrdenado.ToList();
                    _Transfer.Veiculo = new Veiculo(_Transfer.Passageiros.Count);

                    _Transfers.Add(_Transfer);
                }
            }

            return _Transfers;
        }

        #endregion
    }
}