using System;
using System.IO;
using System.Web;
using System.Data;
using System.Linq;
using System.Data.OleDb;
using System.Diagnostics;
using System.ComponentModel;
using System.Collections.Generic;

namespace ChronosTransfer.CLChronos.CLChronosRaiz.CLChronos
{
    [DebuggerNonUserCode(), Browsable(false)]
    public abstract class Conexao : IDisposable
    {
        #region Contrutores

        /// <summary>
        /// Contrutor vazio da conexão que conecta com base na String já existente.
        /// </summary>
        public Conexao() {}

        /// <summary>
        /// Contrutor que renova a origem da conexão baseada no parametro passado.
        /// </summary>
        /// <param name="_Source">Origem do arquivo de conexão (excel)</param>
        public Conexao(String _Source) 
        { 
            Source = _Source; 
        }

        #endregion

        #region Propriedades

        #endregion
        
        public static String Source { get; set; }

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