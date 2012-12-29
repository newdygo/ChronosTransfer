using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace Chronos_Transfer.CLTransfer
{
    public class TransporteVooChegada
    {
        public String NumeroVoo { get; set; }
        public DateTime Data { get; set; }
        public DateTime HorarioSaida { get; set; }
        public DateTime HorarioChegada { get; set; }
        public int Quantidade { get; set; }
        public String TipoVeiculo { get; set; }
    }
}