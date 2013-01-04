using System;
using System.Web;
using System.Data;
using System.Linq;
using System.Data.OleDb;
using System.Globalization;
using System.Web.UI.WebControls;
using System.Collections.Generic;
using ChronosTransfer.CLChronos.CLBasico;
using ChronosTransfer.CLChronos.CLPassageiro;
using ChronosTransfer.CLChronos.CLChronosRaiz.CLChronos;

namespace ChronosTransfer.CLChronos.CLTransfer
{
    public class Transfer : Chronos
    {
        public Transfer(String _Source) : base(_Source) {}

        #region Propriedades
        
        public List<Passageiro> Passageiros { get; set; }
        public Veiculo Veiculo { get; set; }    
        
        #endregion
    }
}