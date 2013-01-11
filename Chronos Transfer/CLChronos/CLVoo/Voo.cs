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
        #region Construtores
        
        public Voo(String _Source) : base(_Source) {}

        #endregion

        #region Propriedades

        /// <summary>
        /// Data do voo.
        /// </summary>
        public DateTime Data { get; set; }    
    
        /// <summary>
        /// Cidade de origem (aeroporto)
        /// </summary>
        public String CidadeOrigem { get; set; }

        /// <summary>
        /// Cidade de destino (aeroporto)
        /// </summary>
        public String CidadeDestino { get; set; }

        /// <summary>
        /// Companhia aérea.
        /// </summary>
        public String CompanhiaAerea { get; set; }

        /// <summary>
        /// Número do voo.
        /// </summary>
        public String NumeroVoo { get; set; }

        /// <summary>
        /// Horário saida.
        /// </summary>
        public DateTime HorarioSaida { get; set; }

        /// <summary>
        /// Horario chegada.
        /// </summary>
        public DateTime HorarioChegada { get; set; }
        
        #endregion
    }
}