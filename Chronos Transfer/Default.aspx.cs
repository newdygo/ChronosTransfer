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
using ChronosTransfer.CLTransfer;
using System.Collections.Generic;

namespace ChronosTransfer
{
    public partial class _Default : Page
    {
        public static String FileName { get; set; }
        static Boolean Pronto = false;

        #region Eventos

        protected void Page_Load(object sender, EventArgs e)
        {
            new Passageiro().Conectar2();
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
            Passageiro _Passageiro = new Passageiro();

            gridSheets.DataSource = _Passageiro.RetornarSchema(OleDbSchemaGuid.Tables_Info);
            gridSheets.DataBind();

            HabilitarComponente(true);            
        }

        /// <summary>
        /// Rotina utilizada para habilitar os botões de comando do upload e processamento.
        /// </summary>
        /// <param name="_Ativa"></param>
        private void HabilitarComponente(Boolean _Ativa)
        {
            btnProcessar.Enabled = _Ativa;
            btnUpload.Enabled = !_Ativa;

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
                    List<VooChegada> _VooChegada = new Passageiro().ProcessarArquivo(FileName, _Row.Cells[1].Text);
                    
                    CreateSheet(_Row.Cells[1].Text, _VooChegada);

                    gridVooChegada.DataSource = _VooChegada;
                    gridVooChegada.DataBind();                                      
                }
            }

            Pronto = true;

            CreateLinkDownload();

            HabilitarComponente(false);            
        }

        /// <summary>
        /// Rotina utilizada para criar uma aba com os resultados de um transfer.
        /// </summary>
        /// <param name="_Sheet"></param>
        /// <param name="_VooChegada"></param>
        private void CreateSheet(String _Sheet, List<VooChegada> _VooChegada)
        {
            Excel _Excel = new Excel();
            _Excel.CreateSheet(_Sheet, _VooChegada);
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