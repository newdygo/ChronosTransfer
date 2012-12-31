using ChronosTransfer.CLTransfer;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace ChronosTransfer
{
    public partial class _Default : Page
    {
        public static String FileName { get; set; }

        protected void btnUpload_Click(object sender, EventArgs e)
        {
            if (FileTransfer.HasFile)
            {
                FileName = Path.GetTempPath() + FileTransfer.FileName;

                FileTransfer.SaveAs(FileName);

                Passageiro _Passageiro = new Passageiro() { DataSource = FileName };

                gridSheets.DataSource = _Passageiro.RetornarSchemaExcel();
                gridSheets.DataBind();

                lblStatusUpload.Text = "Arquivo carregado com sucesso!";

                btnProcessar.Enabled = true;
                btnUpload.Enabled = false;                
            }
        }

        protected void btnProcessar_Click(object sender, EventArgs e)
        {
            gridSheets.Visible = false;
            lblStatusUpload.Visible = false;

            foreach (GridViewRow _Row in gridSheets.Rows)
            {
                CheckBox _CheckBox = (CheckBox)_Row.FindControl("chqBody");

                if (_CheckBox.Checked)
                {
                    Passageiro _Passageiro = new Passageiro() { DataSource = FileName };

                    gridPassageiros.DataSource = _Passageiro.ProcessarArquivo(FileName, _Row.Cells[1].Text);
                    gridPassageiros.DataBind();

                    lblStatusUpload.Text = "Transfer gerado com sucesso.";

                    btnProcessar.Enabled = false;
                    btnUpload.Enabled = true; 

                    break;
                }
            }
        }

        protected void CheckBox1_CheckedChanged(object sender, EventArgs e)
        {
            foreach (GridViewRow _Row in gridSheets.Rows)
            {
                CheckBox _CheckBox = (CheckBox)_Row.FindControl("chqBody");

                if (_CheckBox != null)
                {
                    _CheckBox.Checked = (sender as CheckBox).Checked;
                }
            }
        }
    }
}