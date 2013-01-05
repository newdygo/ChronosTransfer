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

        #region Variáveis

        private String _Nome;
        private Voo _VooChegada;
        private Voo _VooRetorno;

        #endregion

        #region Propriedades

        /// <summary>
        /// Nome do passageiro.
        /// </summary>
        public String Nome 
        {
            get { return _Nome; }
            set 
            {
                if (_Nome == String.Empty)
                {
                    
                }
                else
                {
                    _Nome = value;
                }
            } 
        }

        public String Documento { get; set; }
        public Voo VooChegada { get; set; }
        public Voo VooRetorno { get; set; }

        #endregion       
    }
}