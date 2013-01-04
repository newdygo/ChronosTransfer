using System;
using System.Web;
using System.Data;
using System.Linq;
using System.Data.OleDb;
using System.Globalization;
using System.Web.UI.WebControls;
using System.Collections.Generic;
using ChronosTransfer.CLChronos.CLChronosRaiz.CLChronos;

namespace ChronosTransfer.CLChronos.CLBasico
{
    public class Veiculo : Chronos
    {
        #region Construtores

        public Veiculo() : base() {} 

        public Veiculo(String _Source) : base(_Source) {}

        public Veiculo(int _Capacidade) : base() 
        {
            if (_Capacidade <= 1)
            {
                Nome = "Moto";
                Capacidade = 1;
                ValorViagem = Convert.ToDecimal("150.00");
            }
            else if (_Capacidade <= 4)
            {
                Nome = "Carro";
                Capacidade = 4;
                ValorViagem = Convert.ToDecimal("350.00");
            }
            else if (_Capacidade <= 8)
            {
                Nome = "Mini Van";
                Capacidade = 8;
                ValorViagem = Convert.ToDecimal("700.00");
            }
            else if (_Capacidade <= 22)
            {
                Nome = "Micro Ônibus";
                Capacidade = 22;
                ValorViagem = Convert.ToDecimal("1000.00");
            }
            else if (_Capacidade <= 38)
            {
                Nome = "Ônibus";
                Capacidade = 38;
                ValorViagem = Convert.ToDecimal("1500.00");
            }
        }

        #endregion

        #region Propriedades

        public String Nome { get; set; }
        public int Capacidade { get; set; }
        public Decimal ValorViagem { get; set; }
        
        #endregion
    }
}