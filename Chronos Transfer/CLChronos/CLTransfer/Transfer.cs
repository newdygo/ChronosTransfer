using System;
using System.Web;
using System.Data;
using System.Linq;
using System.Globalization;
using System.Web.UI.WebControls;
using System.Collections.Generic;
using ChronosTransfer.CLChronos.CLVoo;
using ChronosTransfer.CLChronos.CLExcel;
using ChronosTransfer.CLChronos.CLBasico;
using ChronosTransfer.CLChronos.CLPassageiro;
using ChronosTransfer.CLChronos.CLChronosRaiz.CLLog;
using ChronosTransfer.CLChronos.CLChronosRaiz.CLChronos;

namespace ChronosTransfer.CLChronos.CLTransfer
{
    public class Transfer : Chronos
    {
        #region Contrutores
                
        public Transfer(String _Source) : base(_Source) 
        {
            Passageiros = new List<Passageiro>();
            Veiculo = new Veiculo(_Source);
        }

        #endregion

        #region Propriedades

        public String Nome { get; set; }
        public List<Passageiro> Passageiros { get; set; }
        public Veiculo Veiculo { get; set; }    
        
        #endregion

        #region Métodos

        /// <summary>
        /// Rotina utilizada para Gerar os transfers.
        /// </summary>
        /// <param name="_Sheet"></param>
        /// <returns></returns>
        public List<Transfer> GerarTransfers(String _Sheet)
        {
            DataTable _Table = new Excel(Source).RetornarTudoSheet(_Sheet);

            Passageiro _Passageiro = new Passageiro(Source);

            Voo _VooChegada = new Voo(Source);
            Voo _VooRetorno = new Voo(Source);

            Char[] _Char = Environment.NewLine.ToCharArray();

            foreach (DataRow _Row in _Table.Rows)
            {
                try
                {
                    //IFormatProvider _Culture = new System.Globalization.CultureInfo("pt-BR", true);

                    _Passageiro.Nome = _Row[0].ToString();
                    _Passageiro.Documento = _Row[1].ToString();

                    _VooChegada.Data = DateTime.Parse(_Row[2].ToString().Split(_Char).Last().ToString());
                    _VooChegada.CidadeOrigem = _Row[3].ToString().Split(_Char).Last().ToString();
                    _VooChegada.CidadeDestino = _Row[4].ToString().Split(_Char).Last().ToString();
                    _VooChegada.CompanhiaAerea = _Row[5].ToString().Split(_Char).Last().ToString();
                    _VooChegada.NumeroVoo = _Row[6].ToString().Split(_Char).Last().ToString();
                    _VooChegada.HorarioSaida = DateTime.Parse(_Row[7].ToString().Split(_Char).Last().ToString());
                    _VooChegada.HorarioChegada = DateTime.Parse(_Row[8].ToString().Split(_Char).Last().ToString());

                    _VooRetorno.Data = DateTime.Parse(_Row[9].ToString().Split(_Char).First().ToString());
                    _VooRetorno.CidadeOrigem = _Row[10].ToString().Split(_Char).First().ToString();
                    _VooRetorno.CidadeDestino = _Row[11].ToString().Split(_Char).First().ToString();
                    _VooRetorno.CompanhiaAerea = _Row[12].ToString().Split(_Char).First().ToString();
                    _VooRetorno.NumeroVoo = _Row[13].ToString().Split(_Char).First().ToString();
                    _VooRetorno.HorarioSaida = DateTime.Parse(_Row[14].ToString().Split(_Char).First().ToString());
                    _VooRetorno.HorarioChegada = DateTime.Parse(_Row[15].ToString().Split(_Char).First().ToString());

                    _Passageiro.VooChegada = _VooChegada;
                    _Passageiro.VooRetorno = _VooRetorno;

                    Passageiros.Add(_Passageiro);

                    _Passageiro = new Passageiro(Source);
                    _VooChegada = new Voo(Source);
                    _VooRetorno = new Voo(Source);
                }
                catch
                {
                    //Gerar Log do ocorrido e em qual linha e continuar, no final verificar o logo e exibir para o usuário corrigir as pendências.
                    //Não exibir o resultado caso o log seja maior que 0. exibir o log.

                    LogErro _Erro = new LogErro();

                    continue;
                }
            }

            List<Passageiro> _PassageirosOrdenados = Passageiros.OrderBy(x => x.VooChegada.NumeroVoo).ToList();
            List<IGrouping<String, Passageiro>> _PassageirosAgrupados = _PassageirosOrdenados.GroupBy(x => x.VooChegada.NumeroVoo).ToList();

            List<Transfer> _Transfers = new List<Transfer>();

            foreach (IGrouping<String, Passageiro> _PassageiroOrdenado in _PassageirosAgrupados)
            {
                Transfer _Transfer = new Transfer(Source);
                int _CapacidadeMax = Veiculo.GetCapacidadeMaxima();

                if (_PassageiroOrdenado.Count() > _CapacidadeMax)
                {
                    List<Passageiro> _PassageirosTemp = _PassageiroOrdenado.ToList();

                    int _Max = (_PassageirosTemp.Count / _CapacidadeMax);

                    for (int i = 0; i <= _Max; i++)
                    {
                        _Transfer.Nome = String.Format("Transfer {0}.{1}", _PassageirosAgrupados.IndexOf(_PassageiroOrdenado) + 1, i);

                        if (_PassageirosTemp.Count > _CapacidadeMax)
                        {
                            _Transfer.Passageiros = _PassageirosTemp.Take(_CapacidadeMax).ToList();
                        }
                        else
                        {
                            _Transfer.Passageiros = _PassageirosTemp.Take(_PassageirosTemp.Count).ToList();
                        }

                        _Transfer.Veiculo = new Veiculo(Source).CarregarVeiculoCapacidade(_Transfer.Passageiros.Count);

                        _Transfers.Add(_Transfer);

                        _PassageirosTemp.RemoveRange(0, _Transfer.Passageiros.Count);

                        _Transfer = new Transfer(Source);
                    }
                }
                else
                {
                    _Transfer.Nome = String.Format("Transfer {0}", _PassageirosAgrupados.IndexOf(_PassageiroOrdenado) + 1);
                    _Transfer.Passageiros = _PassageiroOrdenado.ToList();
                    _Transfer.Veiculo = new Veiculo(Source).CarregarVeiculoCapacidade(_Transfer.Passageiros.Count);

                    _Transfers.Add(_Transfer);
                }
            }

            return _Transfers;
        }        

        #endregion
    }
}