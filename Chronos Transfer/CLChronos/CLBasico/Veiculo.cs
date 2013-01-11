using System;
using System.Web;
using System.Data;
using System.Linq;
using System.Data.OleDb;
using System.Globalization;
using System.Web.UI.WebControls;
using System.Collections.Generic;
using ChronosTransfer.CLChronos.CLChronosRaiz.CLChronos;
using System.Data.SqlClient;

namespace ChronosTransfer.CLChronos.CLBasico
{
    public class Veiculo : Chronos
    {
        #region Construtores
        
        public Veiculo(String _Source) : base(_Source) {}

        #endregion

        #region Propriedades

        public Int32 IdVeiculo { get; set; }

        /// <summary>
        /// Nome do veículo.
        /// </summary>
        public String Nome { get; set; }

        /// <summary>
        /// Capacidade suportada.
        /// </summary>
        public int Capacidade { get; set; }

        /// <summary>
        /// Valor da viagem.
        /// </summary>
        public Decimal ValorViagem { get; set; }

        /// <summary>
        /// Utilizado na cotação.
        /// </summary>
        public Boolean Utilizado { get; set; }
        
        #endregion

        #region Métodos

        /// <summary>
        /// Rotina utilizada para retornar a capacidade máxima dos carros cadastrados.
        /// </summary>
        /// <returns></returns>
        public int GetCapacidadeMaxima()
        {
            using (ConectarSQL())
            {
                using (SqlCommand _Comand = new SqlCommand("Veiculo_GetCapacidadeMaxima", ConexaoSQL) { CommandType = CommandType.StoredProcedure })
                {
                    return Convert.ToInt32(_Comand.ExecuteScalar());
                }
            }
        }

        /// <summary>
        /// Rotina utilizada para carregar veiculo aproximado pela capacidade passada como parâmetro.
        /// </summary>
        /// <param name="_Capacidade"></param>
        public Veiculo CarregarVeiculoCapacidade(int _Capacidade)
        {
            try
            {
                using (ConectarSQL())
                {
                    using (DataTable _Table = new DataTable())
                    {
                        using (SqlCommand _Comand = new SqlCommand("Veiculo_CarregarVeiculoCapacidade", ConexaoSQL) { CommandType = CommandType.StoredProcedure })
                        {
                            _Comand.Parameters.Add(new SqlParameter("@Capacidade", SqlDbType.Int) { Value = _Capacidade});

                            _Table.Load(_Comand.ExecuteReader());

                            if (_Table.Rows.Count > 0)
                            {
                                IdVeiculo = Convert.ToInt32(_Table.Rows[0][0]);
                                Nome = _Table.Rows[0][1].ToString();
                                Capacidade = Convert.ToInt32(_Table.Rows[0][2]);
                                ValorViagem = Convert.ToInt32(_Table.Rows[0][3]);
                                Utilizado = Convert.ToBoolean(_Table.Rows[0][4]);
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {                
                throw ex;
            }

            return this;
        }

        #endregion
    }
}