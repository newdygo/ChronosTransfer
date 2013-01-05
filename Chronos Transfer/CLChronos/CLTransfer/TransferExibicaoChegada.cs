using System;
using System.Web;
using System.Linq;
using System.Collections.Generic;

namespace ChronosTransfer.CLChronos.CLTransfer
{
    public class TransferExibicaoChegada
    {
        #region Propriedades

        public String Nome { get; set; }
        public String NumeroVoo { get; set; }
        public DateTime Data { get; set; }
        public String CidadeDestino { get; set; }
        public DateTime HorarioSaida { get; set; }
        public DateTime HorarioChegada { get; set; }
        public String Veiculo { get; set; }
        public Decimal Valor { get; set; }

        #endregion

        #region Métodos

        /// <summary>
        /// Rotina utilizada para listar todos os transfer visualmente ao usuário.
        /// </summary>
        /// <param name="_Transfers"></param>
        /// <returns></returns>
        public List<TransferExibicaoChegada> GetExibicaoChegada(List<Transfer> _Transfers)
        {
            List<TransferExibicaoChegada> _TransfersExibicao = new List<TransferExibicaoChegada>();

            foreach (Transfer _Transfer in _Transfers)
            {
                TransferExibicaoChegada _TransferExibicao = new TransferExibicaoChegada();

                _TransferExibicao.Nome = _Transfer.Nome;
                _TransferExibicao.NumeroVoo = _Transfer.Passageiros.First().VooChegada.NumeroVoo;
                _TransferExibicao.Data = _Transfer.Passageiros.First().VooChegada.Data;
                _TransferExibicao.CidadeDestino = _Transfer.Passageiros.First().VooChegada.CidadeDestino;
                _TransferExibicao.HorarioSaida = _Transfer.Passageiros.First().VooChegada.HorarioSaida;
                _TransferExibicao.HorarioChegada = _Transfer.Passageiros.First().VooChegada.HorarioChegada;
                _TransferExibicao.Veiculo = _Transfer.Veiculo.Nome;
                _TransferExibicao.Valor = _Transfer.Veiculo.ValorViagem;

                _TransfersExibicao.Add(_TransferExibicao);
            }

            return _TransfersExibicao;
        }

        #endregion
    }
}