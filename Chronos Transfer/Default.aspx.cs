using Chronos_Transfer.CLTransfer;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

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
                    String _FileName = Path.GetFileName(FileTransfer.FileName);
                    
                    FileTransfer.SaveAs(Server.MapPath("~/") + _FileName);

                    lblStatusUpload.Text = "Upload status: File uploaded!";

                    Passageiro _Passageiro = new Passageiro();

                    _Passageiro.AbrirArquivo(Server.MapPath("~/") + _FileName, gridPassageiros);
                }
                catch (Exception ex)
                {
                    lblStatusUpload.Text = "Upload status: The file could not be uploaded. The following error occured: " + ex.Message;
                }
            }
        }
    }
}