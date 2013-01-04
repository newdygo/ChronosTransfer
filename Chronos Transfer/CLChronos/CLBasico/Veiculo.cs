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
        public Veiculo(String _Source) : base(_Source) { } 

        #region Propriedades

        public String Nome { get; set; }
        public int Capacidade { get; set; }
        public Decimal ValorViagem { get; set; }
        
        #endregion
    }
}