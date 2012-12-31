using System;
using System.IO;
using System.Web;
using System.Data;
using System.Linq;
using System.Data.OleDb;
using System.Collections.Generic;

namespace ChronosTransfer.CLTransfer
{
    public class Excel : Chronos
    {
        public void CreateSheet(String _Sheet, List<VooChegada> _Passageiros)
        {
            using (Conectar())
            {
                using (OleDbCommand command = new OleDbCommand() { CommandType = CommandType.Text, Connection = ConexaoOL })
                {
                    command.CommandText = String.Format("Create Table [Transfer_{0}] ([N° Vôo] varchar(255), [Data] date, [Cidade Origem] varchar(255), [Horário Saída] time, [Horário Chegada] time, [Quantidade Passageiro] int, [Quantidade Veiculo] int, [Veículo] varchar(255))", _Sheet.Replace("$", String.Empty));

                    command.ExecuteNonQuery();
                }
            }

            InsertPassageiroSheet(_Sheet, _Passageiros);
        }

        private void InsertPassageiroSheet(String _Sheet, List<VooChegada> _VoosChegada)
        {
            using (Conectar())
            {
                foreach (VooChegada _Voo in _VoosChegada)
                {
                    using (OleDbCommand command = new OleDbCommand() { CommandType = CommandType.Text, Connection = ConexaoOL })
                    {
                        command.CommandText = String.Format("Insert Into [Transfer_{0}] ([N° Vôo], [Data], [Cidade Origem], [Horário Saída], [Horário Chegada], [Quantidade Passageiro], [Quantidade Veiculo], [Veículo]) values ('{1}', '{2}', '{3}', '{4}', '{5}', '{6}', '{7}', '{8}')", _Sheet.Replace("$", String.Empty), _Voo.NumeroVoo, _Voo.Data, _Voo.CidadeOrigem, _Voo.HorarioSaida, _Voo.HorarioChegada, _Voo.QuantidadePassageiro, _Voo.QuantidadeVeiculo, _Voo.TipoVeiculo);

                        command.ExecuteNonQuery();
                    }
                }
            }
        }
    }
}