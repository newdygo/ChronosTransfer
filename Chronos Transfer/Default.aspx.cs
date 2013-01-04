using System;
using System.IO;
using System.Net;
using System.Web;
using System.Linq;
using System.Data;
using System.Web.UI;
using System.Data.OleDb;
using System.Net.Sockets;
using System.Web.UI.WebControls;
using System.Collections.Generic;
using ChronosTransfer.CLChronos.CLExcel;
using Exc = Microsoft.Office.Interop.Excel;
using ChronosTransfer.CLChronos.CLTransfer;
using ChronosTransfer.CLChronos.CLPassageiro;

namespace ChronosTransfer
{
    public partial class _Default : Page
    {
        public static String FileName { get; set; }
        static Boolean Pronto = false;

        #region Eventos

        /// <summary>
        /// Rotina utilizada quando a página é carregada.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void Page_Load(object sender, EventArgs e)
        {
            CreateLinkDownload();
        }

        /// <summary>
        /// Rotina utilizada no click do botão Upload, que faz a cópia do arquivo local para o servidor e exibe as Sheets que existem no arquivo Excel.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void btnUpload_Click(object sender, EventArgs e)
        {
            CopiarArquivo();
        }

        /// <summary>
        /// Rotina utlizada para processar o arquivo, gerar o link de download.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void btnProcessar_Click(object sender, EventArgs e)
        {
            ProcessarArquivo();
        }

        /// <summary>
        /// Rotina utlizada para cancelar o processo.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void btnCancelar_Click(object sender, EventArgs e)
        {
            HabilitarComponente(false);
        }

        /// <summary>
        /// Rotina utilizada para realizar o download do arquivo Excel modificado já com os resultados do Transfer.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void LinkButton_Click(object sender, EventArgs e)
        {
            Pronto = false;

            FileInfo _File = new FileInfo(FileName);

            if (_File.Exists)
            {
                Response.Clear();
                Response.AddHeader("Content-Disposition", String.Format("attachment; filename = Processado_{0}", _File.Name));

                //Response.AddHeader("Content-Length", _File.Length.ToString());
                //Response.ContentType = "application/octet-stream";

                Response.WriteFile(_File.FullName);
                Response.End();
            }
        }

        /// <summary>
        /// Rotina utilizada para marcar ou desmarcar todos os checkbox do grid quando o header for clicado.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void chqHeader_CheckedChanged(object sender, EventArgs e)
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

        #endregion

        #region Métodos

        /// <summary>
        /// Rotina utilizada para verificar a extensao de um arquivo e copia-lo para pasta temporária do servidor.
        /// </summary>
        private void CopiarArquivo()
        {
            if (FileTransfer.HasFile)
            {
                FileName = String.Format("{0}Transfer.xls", Path.GetTempPath());

                if (Path.GetExtension(FileTransfer.FileName).ToLower() == ".xls")
                {
                    if (File.Exists(FileName))
                    {
                        try
                        {
                            File.Delete(FileName);

                            FileTransfer.SaveAs(FileName);

                            lblStatusUpload.Text = "Arquivo carregado com sucesso!";

                            CarregarSchema();
                        }
                        catch (Exception ex)
                        {
                            lblStatusUpload.Text = String.Format("Arquivo não carregado! \n\nMotivo: {0}", ex.Message);
                        }
                    }
                    else
                    {
                        try
                        {
                            FileTransfer.SaveAs(FileName);

                            lblStatusUpload.Text = "Arquivo carregado com sucesso!";

                            CarregarSchema();
                        }
                        catch (Exception ex)
                        {
                            lblStatusUpload.Text = String.Format("Arquivo não carregado! \n\nMotivo: {0}", ex.Message);
                        }
                    }
                }
            }
            else
            {
                lblStatusUpload.Text = "Arquivo que tentou ser carregado não é um Excel válido!";
            }            
        }

        /// <summary>
        /// Rotina utiliza para carregar o schema de um arquivo excel.
        /// </summary>
        private void CarregarSchema()
        {
            gridSheets.DataSource = new Excel().RetornarSchema(OleDbSchemaGuid.Tables_Info);
            gridSheets.DataBind();
                        
            HabilitarComponente(true);            
        }

        /// <summary>
        /// Rotina utilizada para habilitar os botões de comando do upload e processamento.
        /// </summary>
        /// <param name="_Ativa"></param>
        private void HabilitarComponente(Boolean _Ativa)
        {
            btnProcessar.Visible = _Ativa;
            btnUpload.Visible = !_Ativa;
            
            gridSheets.Visible = _Ativa;
            gridVooChegada.Visible = !_Ativa;

            lblStatusUpload.Visible = !_Ativa;
        }

        private void ProcessarArquivo()
        {
            foreach (GridViewRow _Row in gridSheets.Rows)
            {
                CheckBox _CheckBox = (CheckBox)_Row.FindControl("chqBody");

                if (_CheckBox.Checked)
                {
                    Transfer _Transfer = new Transfer();

                    List<Transfer> _Transfers = _Transfer.GerarTransfers(_Row.Cells[1].Text);

                    Excel _Excel = new Excel();

                    _Excel.CreateSheet(_Row.Cells[1].Text);

                    foreach (Transfer _TransferTemp in _Transfers)
                    {
                        _Excel.InsertPassageiroSheet(_Row.Cells[1].Text, _TransferTemp);
                    }

                    //gridVooChegada.DataSource = _Transfers;
                    //gridVooChegada.DataBind();
                }
            }

            Pronto = true;

            CreateLinkDownload();

            HabilitarComponente(false);            
        }

        /// <summary>
        /// Rotina utilizada para criar o link de download do arquivo gerado pelo transfer.
        /// </summary>
        private void CreateLinkDownload()
        {
            if (Pronto)
            {
                LinkButton _LinkButton = new LinkButton();

                _LinkButton.Text = String.Format("Download Transfer Planilha 'Processado_{0}'", Path.GetFileName(FileName));
                _LinkButton.Click += new EventHandler(LinkButton_Click);

                LinkToDownload.Controls.Add(_LinkButton);
            }
        }

        #endregion
    }
}