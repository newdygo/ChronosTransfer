using System;
using System.Web;
using System.Data;
using System.Linq;
using System.Data.OleDb;
using System.Collections.Generic;
using ChronosTransfer.CLChronos.CLChronosRaiz.CLChronos;

namespace ChronosTransfer.CLTransfer
{
    public class Sheet
    {

        //#region Propriedades

        //public String Nome { get; set; }
        //public Boolean Exibir { get; set; }

        //#endregion

        //#region Métodos

        ///// <summary>
        ///// Rotina utilizada para retornar as Sheets da planilha.
        ///// </summary>
        ///// <param name="_DataSource">Caminho do arquivo de conexão com o Excel.</param>
        ///// <returns>Retorn um datatable com os nomes das sheets da planilha Excel.</returns>
        //public List<Sheet> RetornarSchemaExcel()
        //{
        //    List<Sheet> _Sheets = new List<Sheet>();

        //    using (Conectar())
        //    {
        //        DataTable _Table = ConexaoOL.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);

        //        foreach (DataRow _Row in _Table.Rows)
        //        {
        //            _Sheets.Add(new Sheet() { Nome = _Row[2].ToString(), Exibir = false });
        //        }
        //    }

        //    return _Sheets;
        //}

        //#endregion
    }
}