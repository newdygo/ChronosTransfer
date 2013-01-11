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
        
        public Passageiro(String _Source) : base(_Source) {}

        #endregion
        
        #region Propriedades

        /// <summary>
        /// Nome do passageiro.
        /// </summary>
        public String Nome { get; set; }

        /// <summary>
        /// Documento de identidade.
        /// </summary>
        public String Documento { get; set; }

        /// <summary>
        /// Voo de chegada.
        /// </summary>
        public Voo VooChegada { get; set; }

        /// <summary>
        /// Voo de retorno.
        /// </summary>
        public Voo VooRetorno { get; set; }

        #endregion       
    }
}