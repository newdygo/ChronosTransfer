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
                Passageiro _Passageiro = new Passageiro();

                try
                {
                    String _FileName = Path.GetTempPath() + FileTransfer.FileName;

                    FileTransfer.SaveAs(_FileName);

                    lblStatusUpload.Text = "Upload status: File uploaded!";                    

                    _Passageiro.AbrirArquivo(_FileName, gridPassageiros, gridTeste, ref _Passageiro);
                }
                catch (Exception ex)
                {
                    lblStatusUpload.Text = "Upload status: The file could not be uploaded. The following error occured: " + ex.Message + " " + _Passageiro.Nome;
                }
            }
        }
    }
}