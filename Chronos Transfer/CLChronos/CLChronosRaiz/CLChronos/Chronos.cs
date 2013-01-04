using System;
using System.IO;
using System.Web;
using System.Data;
using System.Linq;
using System.Data.OleDb;
using System.Diagnostics;
using System.ComponentModel;
using System.Collections.Generic;
using ChronosTransfer.CLChronos.CLChronosRaiz.CLLog;

namespace ChronosTransfer.CLChronos.CLChronosRaiz.CLChronos
{
    public abstract class Chronos : Conexao
    {
        #region Construtores

        public Chronos() : base() { }

        public Chronos(String _Source) : base(_Source) {}
        
        #endregion

        #region Propriedades

        public List<LogErro> LogErro { get; set; }

        #endregion        
    }
}