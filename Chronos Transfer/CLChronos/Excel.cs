using System;
using System.IO;
using System.Web;
using System.Data;
using System.Linq;
using System.Data.OleDb;
using System.Collections.Generic;
using ChronosTransfer.CLChronos.CLChronosRaiz.CLChronos;

namespace ChronosTransfer.CLTransfer
{
    public class Excel
    {
        ///// <summary>
        ///// Rotina utilizada para criar uma Sheet num excel.
        ///// </summary>
        ///// <param name="_Sheet"></param>
        ///// <param name="_Passageiros"></param>
        //public void CreateSheet(String _Sheet, List<VooChegada> _Passageiros)
        //{
        //    if (RetornarSchema(OleDbSchemaGuid.Tables_Info, _Sheet) > 0)
        //    {
        //        using (Conectar())
        //        {
        //            using (OleDbCommand _Command = new OleDbCommand() { CommandType = CommandType.Text, Connection = ConexaoOL })
        //            {
        //                _Command.CommandText = String.Format("Drop Table [Transfer_{0}])", _Sheet.Replace("$", String.Empty));
        //                _Command.ExecuteNonQuery();
        //            }
        //        }                
        //    }

        //    using (Conectar())
        //    {
        //        using (OleDbCommand _Command = new OleDbCommand() { CommandType = CommandType.Text, Connection = ConexaoOL })
        //        {   
        //            _Command.CommandText = String.Format("Create Table [Transfer_{0}] ([N° Vôo] int, [Data] date, [Cidade Origem] varchar(255), [Horário Saída] time, [Horário Chegada] time, [Quantidade Passageiro] int, [Quantidade Veiculo] int, [Veículo] varchar(255))", _Sheet.Replace("$", String.Empty));
        //            _Command.ExecuteNonQuery();
        //        }
        //    }

        //    InsertPassageiroSheet(_Sheet, _Passageiros);
        //}

        ///// <summary>
        ///// Rotina utilizada para retornar todos os dados de uma sheet
        ///// </summary>
        ///// <param name="_Sheet">Sheet da planilha Excel.</param>
        ///// <returns>Retorna dados de uma sheet Excel.</returns>
        //public DataTable RetornarTudoSheet(String _Sheet)
        //{
        //    using (Conectar())
        //    {
        //        //No momento da abertura do arquivo identificar quais planilhas existem dentro da pasta e exibir ao usuário em formato de combobox o nome das planilhas
        //        //e pedir que identifique qual planilha contem os dados para processamento.

        //        using (OleDbCommand _Command = new OleDbCommand() { CommandText = String.Format("Select * From [{0}]", _Sheet), CommandType = CommandType.Text, Connection = ConexaoOL })
        //        {
        //            using (DataTable _Table = new DataTable())
        //            {
        //                _Table.Load(_Command.ExecuteReader());

        //                return _Table;
        //            }
        //        }
        //    }
        //}

        ///// <summary>
        ///// Rotina utilizada para inserir voo de chegada a sheet do Excel.
        ///// </summary>
        ///// <param name="_Sheet"></param>
        ///// <param name="_VoosChegada"></param>
        //private void InsertPassageiroSheet(String _Sheet, List<VooChegada> _VoosChegada)
        //{
        //    using (Conectar())
        //    {
        //        foreach (VooChegada _Voo in _VoosChegada)
        //        {
        //            using (OleDbCommand _Command = new OleDbCommand() { CommandType = CommandType.Text, Connection = ConexaoOL })
        //            {
        //                _Command.CommandText = String.Format("Insert Into [Transfer_{0}] ([N° Vôo], [Data], [Cidade Origem], [Horário Saída], [Horário Chegada], [Quantidade Passageiro], [Quantidade Veiculo], [Veículo]) values ('{1}', '{2}', '{3}', '{4}', '{5}', '{6}', '{7}', '{8}')", _Sheet.Replace("$", String.Empty), _Voo.NumeroVoo, _Voo.Data, _Voo.CidadeOrigem, _Voo.HorarioSaida, _Voo.HorarioChegada, _Voo.QuantidadePassageiro, _Voo.QuantidadeVeiculo, _Voo.TipoVeiculo);

        //                _Command.ExecuteNonQuery();
        //            }
        //        }
        //    }
        //}
    }
}