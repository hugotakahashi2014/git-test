using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using VSLibrary;
using Microsoft.Win32;
using mshtml;

namespace ConsultaCadastro
{
    public partial class FormConsultaCCCRFB : BaseForm
    {

        private string _cnpj = string.Empty;
        //private string _ie = string.Empty;
        private TipoDadosCadastroCNPJ _Registro;
        private string[] _tiposLograd = null;
        //setando url do CCC - Cadastro Centralizado de Contribuintes
        private const string _url = "https://dfe-portal.svrs.rs.gov.br/NFE/CCC";

        private TipoDadosCadastroCNPJ Registro
        {
            get { return _Registro; }
        }

        public FormConsultaCCCRFB()
        {

            InitializeComponent();
            VerificaVersaoIE();

        }

        private void VerificaVersaoIE()
        {
            //esta funcao contorna um bug do WebBrowser
            string appName = System.Diagnostics.Process.GetCurrentProcess().ProcessName + ".exe";

            RegistryKey fbeKey = null;
            ////For 64 bit Machine 
            fbeKey = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(@"SOFTWARE\\Wow6432Node\\Microsoft\\Internet Explorer\\MAIN\\FeatureControl\\FEATURE_BROWSER_EMULATION", true);
            if (fbeKey == null)
                fbeKey = Registry.CurrentUser.CreateSubKey(@"SOFTWARE\\Wow6432Node\\Microsoft\\Internet Explorer\\MAIN\\FeatureControl\\FEATURE_BROWSER_EMULATION");
            using (fbeKey)
            {
                fbeKey.SetValue(appName, 11001, RegistryValueKind.DWord);
            }


            //For 32 bit Machine 
            fbeKey = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(@"SOFTWARE\\Microsoft\\Internet Explorer\\Main\\FeatureControl\\FEATURE_BROWSER_EMULATION", true);
            if (fbeKey == null)
                fbeKey = Registry.CurrentUser.CreateSubKey(@"SOFTWARE\\Microsoft\\Internet Explorer\\Main\\FeatureControl\\FEATURE_BROWSER_EMULATION");
            using (fbeKey)
            {
                fbeKey.SetValue(appName, 11001, RegistryValueKind.DWord);
            }
        }

        public DialogResult CarregaPagina(string Cnpj, string[] TiposLogradouro, Form FormPai)
        {
            _cnpj = Cnpj;
            _tiposLograd = TiposLogradouro;
            _Registro = new TipoDadosCadastroCNPJ();

            return this.ExibeDialog(FormPai);
        }

        private void FormConsultaCCCRFB_AoExibirForm(object Sender, EventArgs e)
        {
            this.DialogResult = DialogResult.Cancel;
            cmdAtualizar.PerformClick();
        }

        private void cmdAtualizar_Click(object sender, EventArgs e)
        {
            try
            {
                cmdAtualizar.Enabled = false;
                cmdCampos.Enabled = false;

                webBrowser1.Navigate(_url);
                webBrowser1.Navigating += new WebBrowserNavigatingEventHandler(webBrowser1_Navigating);
            }
            finally
            {
                cmdAtualizar.Enabled = true;
                cmdCampos.Enabled = true;
            }
        }
        private void webBrowser1_Navigating(object sender, WebBrowserNavigatingEventArgs e)
        {
            HtmlDocument document = this.webBrowser1.Document;
            //HtmlElement elementCpfCnpj = document.GetElementsByTagName("CodInscrMf")

            if (document == null)
            {
                e.Cancel = true;
                MessageBox.Show(string.Format("O Documento está vazio em {0}", e.Url.ToString()));
            }

            if (document.All["CodInscrMf"] == null)
            {
                e.Cancel = true;
                MessageBox.Show(string.Format("O documento em {0} não contém a tag de Cpf/Cnpj", e.Url.ToString()));
            }

            if (String.IsNullOrEmpty(document.All["CodInscrMf"].GetAttribute("value")))
            {
                //preencher valor do campo javascript
                e.Cancel = true;
                MessageBox.Show(string.Format("cpf/cnpj {0}", document.All["CodInscrMf"].GetAttribute("value").ToString()));

            }

            HtmlElement elementHead = webBrowser1.Document.GetElementsByTagName("head")[0];
            HtmlElement elementScript = webBrowser1.Document.CreateElement("script");
            IHTMLScriptElement DOMScriptElement = (IHTMLScriptElement)elementScript.DomElement;
            DOMScriptElement.text = string.Format("function SetarCNPJ(){" +
                                                    "document.getElementById('CodInscrMf').value =  {0}", _cnpj);

            elementHead.AppendChild(elementScript);
            webBrowser1.Document.InvokeScript("SetarCNPJ");
   
        }
    }
}
