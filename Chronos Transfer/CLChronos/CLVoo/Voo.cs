using System;
using System.Web;
using System.Data;
using System.Linq;
using System.Data.OleDb;
using System.Globalization;
using System.Web.UI.WebControls;
using System.Collections.Generic;
using ChronosTransfer.CLChronos.CLChronosRaiz.CLChronos;

namespace ChronosTransfer.CLChronos.CLVoo
{
    public class Voo : Chronos
    {
        public Voo(String _Source) : base(_Source) {}

        #region Propriedades
        
        public DateTime Data { get; set; }
        public String NumeroVoo { get; set; }
        public String CidadeOrigem { get; set; }
        public String CidadeDestino { get; set; }
        public String CompanhiaAerea { get; set; }        
        public DateTime HorarioSaida { get; set; }
        public DateTime HorarioChegada { get; set; }
        
        #endregion
    }
}