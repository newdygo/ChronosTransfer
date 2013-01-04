using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace ChronosTransfer.CLTransfer
{
    public class VooChegada
    {
        public String NumeroVoo { get; set; }
        public DateTime Data { get; set; }
        public String CidadeOrigem { get; set; }
        public DateTime HorarioSaida { get; set; }
        public DateTime HorarioChegada { get; set; }
        public int QuantidadePassageiro { get; set; }
        public int QuantidadeVeiculo { get; set; }
        public String TipoVeiculo { get; set; }
    }
}