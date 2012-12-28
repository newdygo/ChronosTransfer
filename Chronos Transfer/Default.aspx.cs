using Chronos_Transfer.CLTransfer;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using Excel = Microsoft.Office.Interop.Excel;

namespace Chronos_Transfer
{
    public partial class _Default : Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {

        }

        protected void btnProcessar_Click(object sender, EventArgs e)
        {
            if (FileTransfer.HasFile)
            {
                try
                {
                    FileTransfer.SaveAs(Server.MapPath("~/") + FileTransfer.FileName);

                    Excel.Application excel = new Excel.Application();
                    Excel.Workbook wb = excel.Workbooks.Open(Server.MapPath("~/") + FileTransfer.FileName);

                    foreach (Excel.Worksheet sh in wb.Worksheets)
                    {
                        lblStatusUpload.Text = sh.Cells[1, "A"].Value2;
                        break;
                    }

                    wb.Close();
                    excel.Quit();
                     

                    //String _FileName = Path.GetFileName(FileTransfer.FileName);

                    //String jjj = Path.GetFullPath(FileTransfer.FileName);
                    
                    //lblStatusUpload.Text = "Upload status: File uploaded!";

                    //Passageiro _Passageiro = new Passageiro();

                    //_Passageiro.AbrirArquivo(jjj, gridPassageiros);
                }
                catch (Exception ex)
                {
                    lblStatusUpload.Text = "Upload status: The file could not be uploaded. The following error occured: " + ex.Message;
                }
            }
        }
    }
}