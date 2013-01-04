﻿using System;
using System.IO;
using System.Web;
using System.Data;
using System.Linq;
using System.Data.OleDb;
using System.Collections.Generic;
using ChronosTransfer.CLChronos.CLChronosRaiz.CLChronos;
using ChronosTransfer.CLChronos.CLTransfer;
using ChronosTransfer.CLChronos.CLPassageiro;

namespace ChronosTransfer.CLChronos.CLExcel
{
    public class Excel : Chronos
    {
        #region Construtores

        public Excel() : base() { }

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
            using (Conectar())
            {
                using (OleDbCommand _Command = new OleDbCommand() { CommandText = String.Format("Select * From [{0}]", _Sheet), CommandType = CommandType.Text, Connection = ConexaoOL })
                {
                    using (DataTable _Table = new DataTable())
                    {
                        _Table.Load(_Command.ExecuteReader());

                        return _Table;
                    }
                }
            }
        }

        #endregion

        #region Úteis

        /// <summary>
        /// Rotina utilizada para retornar as Sheets da planilha.
        /// </summary>
        /// <param name="_DataSource">Caminho do arquivo de conexão com o Excel.</param>
        /// <returns>Retorn um datatable com os nomes das sheets da planilha Excel.</returns>
        public DataTable RetornarSchema(Guid _OleDbSchemaGuid)
        {
            using (Conectar())
            {
                return ConexaoOL.GetOleDbSchemaTable(_OleDbSchemaGuid, null).Select("Cardinality = 0").CopyToDataTable();
            }
        }

        /// <summary>
        /// Rotina utilizada para retornar as Sheets da planilha.
        /// </summary>
        /// <param name="_DataSource">Caminho do arquivo de conexão com o Excel.</param>
        /// <returns>Retorn um datatable com os nomes das sheets da planilha Excel.</returns>
        public int RetornarSchema(Guid _OleDbSchemaGuid, String _Sheet)
        {
            using (Conectar())
            {
                return ConexaoOL.GetOleDbSchemaTable(_OleDbSchemaGuid, null).Select(String.Format("Cardinality = 0 And Table_Name = 'Transfer_{0}'", _Sheet)).Count();
            }
        }

        /// <summary>
        /// Rotina utilizada para retornar as Sheets da planilha.
        /// </summary>
        /// <param name="_DataSource">Caminho do arquivo de conexão com o Excel.</param>
        /// <returns>Retorn um datatable com os nomes das sheets da planilha Excel.</returns>
        public DataTable RetornarColumnsSchema(Guid _OleDbSchemaGuid, String _Sheet)
        {
            using (Conectar())
            {
                return ConexaoOL.GetOleDbSchemaTable(_OleDbSchemaGuid, null);
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
            if (RetornarSchema(OleDbSchemaGuid.Tables_Info, _Sheet) > 0)
            {
                using (Conectar())
                {
                    using (OleDbCommand _Command = new OleDbCommand() { CommandType = CommandType.Text, Connection = ConexaoOL })
                    {
                        _Command.CommandText = String.Format("Drop Table [Transfer_{0}])", _Sheet.Replace("$", String.Empty));
                        _Command.ExecuteNonQuery();
                    }
                }
            }

            using (Conectar())
            {
                using (OleDbCommand _Command = new OleDbCommand() { CommandType = CommandType.Text, Connection = ConexaoOL })
                {
                    _Command.CommandText = String.Format("Create Table [Transfer_{0}] ([Nome] varchar(255), [Documento] varchar(255), [N° Vôo] varchar(10), [Data] date, [Cidade Origem] varchar(255), [Cidade Destino] varchar(255), [Horário Saída] time, [Horário Chegada] time, [Veículo] varchar(255), [Valor] float)", _Sheet.Replace("$", String.Empty));
                    _Command.ExecuteNonQuery();
                }
            }
        }

        /// <summary>
        /// Rotina utilizada para inserir voo de chegada a sheet do Excel.
        /// </summary>
        /// <param name="_Sheet"></param>
        /// <param name="_Transfer"></param>
        public void InsertPassageiroSheet(String _Sheet, Transfer _Transfer)
        {
            using (Conectar())
            {
                foreach (Passageiro _Passageiro in _Transfer.Passageiros)
                {
                    using (OleDbCommand _Command = new OleDbCommand() { CommandType = CommandType.Text, Connection = ConexaoOL })
                    {
                        _Command.CommandText = String.Format("Insert Into [Transfer_{0}] ([Nome], [Documento], [N° Vôo], [Data], [Cidade Origem], [Cidade Destino], [Horário Saída], [Horário Chegada], [Veículo], [Valor]) values ('{1}', '{2}', '{3}', '{4}', '{5}', '{6}', '{7}', '{8}', '{9}', '{10}')",
                            _Sheet.Replace("$", String.Empty),
                            _Passageiro.Nome, _Passageiro.Documento, _Passageiro.VooChegada.NumeroVoo, _Passageiro.VooChegada.Data,
                            _Passageiro.VooChegada.CidadeOrigem, _Passageiro.VooChegada.CidadeDestino, _Passageiro.VooChegada.HorarioSaida,
                            _Passageiro.VooChegada.HorarioChegada, _Transfer.Veiculo.Nome, _Transfer.Veiculo.ValorViagem);

                        _Command.ExecuteNonQuery();
                    }
                }   
            }
        }
    }
}