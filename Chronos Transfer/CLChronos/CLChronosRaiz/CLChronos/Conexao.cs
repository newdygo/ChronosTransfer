using System;
using System.IO;
using System.Web;
using System.Data;
using System.Linq;
using OfficeOpenXml;
using System.Data.OleDb;
using System.Diagnostics;
using System.Configuration;
using System.Data.SqlClient;
using System.ComponentModel;
using System.Collections.Generic;

namespace ChronosTransfer.CLChronos.CLChronosRaiz.CLChronos
{
    //[DebuggerNonUserCode(), Browsable(false)]
    public abstract class Conexao
    {
        #region Contrutores
        
        /// <summary>
        /// Contrutor que renova a origem da conexão baseada no parametro passado.
        /// </summary>
        /// <param name="_Source">Origem do arquivo de conexão (excel)</param>
        public Conexao(String _Source) 
        { 
            Source = _Source;

            ConexaoSQL = new SqlConnection();
        }

        #endregion

        #region Propriedades

        public ExcelPackage ConexaoEx { get; set; }
        public SqlConnection ConexaoSQL { get; set; }
        public static String Source { get; set; }

        #endregion
        
        #region Métodos
        
        /// <summary>
        /// Rotina utilizada para conectar a um banco de dados do tipo Excel.
        /// </summary>
        public ExcelPackage ConectarEx()
        {
            try
            {
                ConexaoEx = new ExcelPackage(new FileInfo(Source));
            }
            catch (Exception)
            {
                throw new Exception("Excel indisponível.");
            }

            return ConexaoEx;
        }

        /// <summary>
        /// Rotina utilizada para conectar a um banco de dados do SQL Server.
        /// </summary>
        /// <returns></returns>
        public SqlConnection ConectarSQL()
        {
            if (ConexaoSQL.State != ConnectionState.Open)
            {
                try
                {
                    String _StringConexao = ConfigurationManager.ConnectionStrings["SQLConnectionString"].ConnectionString;

                    if (_StringConexao != null)
                    {
                        ConexaoSQL = new SqlConnection(_StringConexao);

                        ConexaoSQL.Open();
                    }
                }
                catch (Exception)
                {
                    throw new Exception("Banco de dados indisponível no momento.");
                }
            }

            return ConexaoSQL;
        }

        #endregion
    }
}