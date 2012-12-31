using ChronosTransfer.CLTransfer;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Sockets;
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

                gridSheets.DataSource = _Passageiro.RetornarSchema(OleDbSchemaGuid.Tables);
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

                    gridVooChegada.DataSource = _Passageiro.ProcessarArquivo(FileName, _Row.Cells[1].Text);
                    gridVooChegada.DataBind();

                    Excel _Excel = new Excel() { DataSource = FileName };
                    _Excel.CreateSheet(_Row.Cells[1].Text, (List<VooChegada>)gridVooChegada.DataSource);
                    
                    lblStatusUpload.Text = "Transfer gerado com sucesso.";

                    CreateLinkDownload();

                    btnProcessar.Enabled = false;
                    btnUpload.Enabled = true; 

                    break;
                }
            }
        }

        private void CreateLinkDownload()
        {
            //string path = Server.MapPath(filename);

            System.IO.FileInfo file = new System.IO.FileInfo(FileName);
            if (file.Exists)
            {
                Response.Clear();
                Response.AddHeader("Content-Disposition", "attachment; filename=" + file.Name);
                Response.AddHeader("Content-Length", file.Length.ToString());
                Response.ContentType = "application/octet-stream";
                Response.WriteFile(file.FullName);
                Response.End();
            }
            else
            {
                Response.Write("This file does not exist.");
            }

            //String jjj = HttpContext.Current.Server.MapPath(FileName);
            
            //IPHostEntry host;
            //string localIP = "?";
            //object diego = Dns.GetHostEntry(Dns.GetHostName()).AddressList.Where(x => x.AddressFamily == AddressFamily.InterNetwork);

            //foreach (IPAddress ip in host.AddressList)
            //{
            //    if (ip.AddressFamily.ToString() == "InterNetwork")
            //    {
            //        localIP = ip.ToString();
            //    }
            //}

            //IPAddress[] host;
            //host = Dns.GetHostAddresses(System.Environment.MachineName);
            //string ip = host[0].ToString() + "\\" + Path.GetTempPath();

            LinkToDownload.Controls.Add(new LinkButton() { Text = String.Format("Download Transfer Planilha 'Transfer_{0}'", FileName), PostBackUrl = FileName });
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