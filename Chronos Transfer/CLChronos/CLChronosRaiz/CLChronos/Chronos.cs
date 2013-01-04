using System;
using System.Web;
using System.Linq;
using System.Collections.Generic;
using ChronosTransfer.CLChronos.CLChronosRaiz.CLLog;

namespace ChronosTransfer.CLChronos.CLChronosRaiz.CLChronos
{
    public abstract class Chronos : Conexao
    {
        public Chronos(String _Source) : base(_Source) {}

        List<LogErro> LogErro { get; set; }
    }
}