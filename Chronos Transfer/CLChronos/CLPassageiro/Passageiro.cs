using System;
using System.Web;
using System.Data;
using System.Linq;
using System.Data.OleDb;
using System.Globalization;
using System.Web.UI.WebControls;
using System.Collections.Generic;
using ChronosTransfer.CLChronos.CLVoo;
using ChronosTransfer.CLChronos.CLChronosRaiz.CLChronos;

namespace ChronosTransfer.CLChronos.CLPassageiro
{
    public class Passageiro : Chronos
    {
        #region Construtores

        public Passageiro() : base() {}

        public Passageiro(String _Source) : base(_Source) {}

        #endregion

        #region Propriedades

        public String Nome { get; set; }
        public String Documento { get; set; }
        public Voo VooChegada { get; set; }
        public Voo VooRetorno { get; set; }

        #endregion       
    }
}