using System;
using System.Web;
using System.Data;
using System.Linq;
using System.Data.OleDb;
using System.Diagnostics;
using System.ComponentModel;
using System.Collections.Generic;
using System.IO;

namespace ChronosTransfer.CLTransfer
{
    [DebuggerNonUserCode(), Browsable(false)]
    public abstract class Conexao : IDisposable
    {
        #region Variáveis

        public OleDbConnection ConexaoOL = new OleDbConnection();
        
        #endregion

        #region OleDbConnextion

        /// <summary>
        /// Rotina utilizada para conectar a um banco de dados
        /// </summary>
        public OleDbConnection Conectar()
        {
            if (ConexaoOL.State != ConnectionState.Open)
            {
                try
                {
                    //Conexão OleDb para versões atuais do excel, não funciona na WEB (necessita descobrir o motivo)
                    //String _Conexao = String.Format("Provider=Microsoft.Ace.OleDb.12.0; Data Source= {0}; Extended Properties=\"Excel 12.0 Xml; HDR=Yes\";", _DataSource);

                    String _Conexao = String.Format("Provider=Microsoft.Jet.OLEDB.4.0; Data Source= {0}Transfer.xls; Extended Properties=\"Excel 8.0; HDR=Yes\";", Path.GetTempPath());

                    ConexaoOL = new OleDbConnection(_Conexao);

                    ConexaoOL.Open();
                }
                catch (Exception)
                {                    
                    throw new Exception("Excel indisponível.");
                }                
            }

            return ConexaoOL;
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
                using (DataTable _Table = new DataTable())
                {
                    return ConexaoOL.GetOleDbSchemaTable(_OleDbSchemaGuid, null);
                }
            }
        }

        #endregion

        #region IDisposable Members

        public void Dispose()
        {
            if (ConexaoOL != null)
            {
                ConexaoOL.Dispose();
            }
        }

        #endregion
    }
}