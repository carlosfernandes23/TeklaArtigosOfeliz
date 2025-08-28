using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Net.Mail;
using System.Net;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.Tab;
using CefSharp.WinForms;
using CefSharp;
using System.Net.Mime;
using System.IO;
using Tekla.Structures.Model;
using Newtonsoft.Json;
using HtmlAgilityPack;
using HtmlDocument = HtmlAgilityPack.HtmlDocument;
using Microsoft.Web.WebView2.WinForms;
using Microsoft.Web.WebView2.Core;
using System.Net.Http;
using Microsoft.Identity.Client;
using Tekla.Structures.InpParser;
using Guna.Charts.WinForms;
using DocumentFormat.OpenXml.Office2016.Drawing.ChartDrawing;
using MimeKit;

namespace TeklaArtigosOfeliz
{
    public partial class Frm_Corpo_de_Texto_Email_Pesclarecimento : Form
    {
        private string CorpoEmailImagens;
        private string Subject;

        public Frm_Corpo_de_Texto_Email_Pesclarecimento(string titulo, string corpoEmail, string Subject)
        {
            InitializeComponent();
            this.Text = titulo;
            this.CorpoEmailImagens = corpoEmail;
            this.Subject = Subject;
        }

        private void Frm_Corpo_de_Texto_Email_Pesclarecimento_Load(object sender, EventArgs e)
        {
            textBoxAsu.Text = Subject;
            webBrowser1.DocumentText = CorpoEmailImagens;
            CarregarDiretorObra();
            CarregarEmailsCC();
            VerificarUser();
            string user = Environment.UserName;
            labelemailuser.Text = user + "@ofeliz.com";
            Model modelo = new Model();
            string numeroobra = modelo.GetProjectInfo().ProjectNumber;
            label11.Text = numeroobra;
        }

        private List<string> LoadNamesFromJson(string filePath)
        {
            try
            {
                string json = File.ReadAllText(filePath);

                List<string> nomes = JsonConvert.DeserializeObject<List<string>>(json);

                return nomes;
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, "Erro ao carregar o arquivo JSON: " + ex.Message);
                return new List<string>();
            }
        }

        private void CarregarEmailsCC()
        {
            string caminho = @"\\marconi\COMPANY SHARED FOLDER\OFELIZ\OFM\2.AN\2.CM\DP\4 Produção\Desenvolvimentos\Ficheiros Temp tekla artigos (Nao Apagar)\Diretor de Obra Base de dados\EmailDuvidasPECC.json";

            try
            {
                if (File.Exists(caminho))
                {
                    string json = File.ReadAllText(caminho);

                    var emails = JsonConvert.DeserializeObject<List<string>>(json);

                    listBoxCC.Items.Clear();
                    foreach (var email in emails)
                    {
                        listBoxCC.Items.Add(email);
                    }
                }
                else
                {
                    MessageBox.Show("Ficheiro dos email CC Não encontrado.");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Erro ao carregar ficheiro: " + ex.Message);
            }
        }

        private void CarregarDiretorObra()
        {
            comboBoxDiretorObra.Items.Clear();
            string jsonFilePath = @"\\marconi\COMPANY SHARED FOLDER\OFELIZ\OFM\2.AN\2.CM\DP\4 Produção\Desenvolvimentos\Ficheiros Temp tekla artigos (Nao Apagar)\Diretor de Obra Base de dados\DiretordeObra.json";
            List<string> nomes = LoadNamesFromJson(jsonFilePath);
            foreach (var nome in nomes)
            {
                comboBoxDiretorObra.Items.Add(nome);
            }
        }

        private void VerificarUser()
        {
            string user = Environment.UserName.ToLower();

            if (user == "luis.silva")
            {
                listBoxCC.Items.Remove("luis.silva@ofeliz.com");
            }
            else if (user == "helder.silva")
            {
                listBoxCC.Items.Remove("helder.silva@ofeliz.com");
            }

        }

        private void comboBoxDiretorObra_SelectedIndexChanged(object sender, EventArgs e)
        {
            string nomeSelecionado = comboBoxDiretorObra.Text.Trim();

            if (!string.IsNullOrEmpty(nomeSelecionado))
            {
                string emailFormatado = nomeSelecionado.ToLower().Replace(" ", ".") + "@ofeliz.com";

                if (!listBoxPara.Items.Contains(emailFormatado))
                {
                    listBoxPara.Items.Add(emailFormatado);
                }
            }
        }

        private void guna2Button4_Click(object sender, EventArgs e)
        {
            this.Visible = false;
            Frm_AtualizarEmails F = new Frm_AtualizarEmails();
            F.ShowDialog();
            this.Visible = true;
        }

        private void guna2Button5_Click(object sender, EventArgs e)
        {
            CarregarDiretorObra();

        }

        private void listBoxCC_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Delete && listBoxCC.SelectedItem != null)
            {
                listBoxCC.Items.Remove(listBoxCC.SelectedItem);
            }

            if (e.Control && e.KeyCode == Keys.C && listBoxCC.SelectedItem != null)
            {
                Clipboard.SetText(listBoxCC.SelectedItem.ToString());
            }

            if (e.Control && e.KeyCode == Keys.V)
            {
                string textoCopiado = Clipboard.GetText();
                if (!string.IsNullOrWhiteSpace(textoCopiado))
                {
                    if (!listBoxCC.Items.Contains(textoCopiado))
                    {
                        listBoxCC.Items.Add(textoCopiado);
                    }
                }
            }
        }

        private void listBoxPara_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Delete && listBoxPara.SelectedItem != null)
            {
                listBoxPara.Items.Remove(listBoxPara.SelectedItem);
            }

            if (e.Control && e.KeyCode == Keys.C && listBoxPara.SelectedItem != null)
            {
                Clipboard.SetText(listBoxPara.SelectedItem.ToString());
            }

            if (e.Control && e.KeyCode == Keys.V)
            {
                string textoCopiado = Clipboard.GetText();
                if (!string.IsNullOrWhiteSpace(textoCopiado))
                {
                    if (!listBoxPara.Items.Contains(textoCopiado))
                    {
                        listBoxPara.Items.Add(textoCopiado);
                    }
                }
            }
        }

        Dictionary<string, string> mimeTypes = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
                {
                    { ".pdf", "application/pdf" },
                    { ".xls", "application/vnd.ms-excel" },
                    { ".xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" },
                    { ".xlsm", "application/vnd.ms-excel.sheet.macroEnabled.12" },
                    { ".doc", "application/msword" },
                    { ".docx", "application/vnd.openxmlformats-officedocument.wordprocessingml.document" },
                    { ".ppt", "application/vnd.ms-powerpoint" },
                    { ".pptx", "application/vnd.openxmlformats-officedocument.presentationml.presentation" },
                    { ".txt", "text/plain" },
                    { ".rtf", "application/rtf" },
                    { ".csv", "text/csv" },
                    { ".jpg", "image/jpeg" },
                    { ".jpeg", "image/jpeg" },
                    { ".png", "image/png" },
                    { ".gif", "image/gif" },
                    { ".bmp", "image/bmp" },
                    { ".dwg", "image/vnd.dwg" },
                    { ".zip", "application/zip" },
                    { ".rar", "application/vnd.rar" },
                    { ".7z", "application/x-7z-compressed" },
                    { ".mp3", "audio/mpeg" },
                    { ".mp4", "video/mp4" },
                    { ".avi", "video/x-msvideo" },
                    { ".mov", "video/quicktime" },
                    { ".html", "text/html" },
                    { ".htm", "text/html" },
                    { ".xml", "application/xml" },
                    { ".ifc", "application/x-step" }
                };

        private Dictionary<string, string> ficheirosSelecionados = new Dictionary<string, string>();

        private string GetMimeType(string extensao)
        {
            return mimeTypes.TryGetValue(extensao.ToLower(), out string mime)
                ? mime
                : "application/octet-stream";
        }

        private void listBoxFicheiros_DragDrop(object sender, DragEventArgs e)
        {
            string[] arquivosArrastados = (string[])e.Data.GetData(DataFormats.FileDrop);

            foreach (string caminhoArquivo in arquivosArrastados)
            {
                string extensao = Path.GetExtension(caminhoArquivo).ToLower();

                if (mimeTypes.ContainsKey(extensao))
                {
                    string nomeArquivo = Path.GetFileName(caminhoArquivo);

                    if (!listBoxFicheiros.Items.Contains(nomeArquivo))
                    {
                        listBoxFicheiros.Items.Add(nomeArquivo);
                        ficheirosSelecionados[nomeArquivo] = caminhoArquivo;
                    }
                }
                else
                {
                    MessageBox.Show("Tipo de arquivo não suportado: " + extensao, "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void listBoxFicheiros_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                e.Effect = DragDropEffects.Copy;
            }
            else
            {
                e.Effect = DragDropEffects.None;
            }
        }

        private void listBoxFicheiros_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Delete && listBoxFicheiros.SelectedItem != null)
            {
                listBoxFicheiros.Items.Remove(listBoxFicheiros.SelectedItem);
            }
        }

        private async void Buttonenviar_Click(object sender, EventArgs e)
        {
            try
            {
                if (listBoxPara.Items.Count > 0)
                {
                    await SendEmailAsync(textBoxAsu);
                }
                else
                {
                    CustomGradientPanel4.Visible = true;
                    CustomGradientPanel4.FillColor = Color.Maroon;
                    CustomGradientPanel4.FillColor2 = Color.Maroon;
                    CustomGradientPanel4.FillColor3 = Color.Maroon;
                    CustomGradientPanel4.FillColor4 = Color.Maroon;

                    labelEmailInformação.ForeColor = Color.White;
                    labelEmailInformação.Location = new Point(10, 8);
                    labelEmailInformação.Text = "Por favor, insira o nome do diretor de obra.";
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, $"Erro: {ex.Message}", "Erro ao enviar e-mail", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private string GetSaudacao()
        {
            DateTime horaAtual = DateTime.Now;
            if (horaAtual.Hour < 12 || (horaAtual.Hour == 12 && horaAtual.Minute < 30))
            {
                return "Bom Dia, ";
            }
            else
            {
                return "Boa Tarde, ";
            }
        }

        public static async Task<string> GetAccessTokenAsync()
        {
            var clientId = "0f37e406-80fc-4deb-9635-a20c4a22c53e";
            var tenantId = "67345170-c562-4f1e-aef4-cf8d2d06067f";
            var redirectUri = "http://localhost:61658";

            var publicClientApplication = PublicClientApplicationBuilder.Create(clientId)
                .WithAuthority(AzureCloudInstance.AzurePublic, tenantId)
                .WithRedirectUri(redirectUri)
                .Build();

            var tokenCache = publicClientApplication.UserTokenCache;
            string cacheFilePath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "msal_cache.dat");

            tokenCache.SetBeforeAccess(args =>
            {
                args.TokenCache.DeserializeMsalV3(File.Exists(cacheFilePath) ? File.ReadAllBytes(cacheFilePath) : null);
            });

            tokenCache.SetAfterAccess(args =>
            {
                if (args.HasStateChanged)
                {
                    File.WriteAllBytes(cacheFilePath, args.TokenCache.SerializeMsalV3());
                }
            });

            var accounts = await publicClientApplication.GetAccountsAsync();
            var account = accounts.FirstOrDefault();

            if (account == null)
            {
                var result = await publicClientApplication
                    .AcquireTokenInteractive(new[] { "Mail.Send" })
                    .ExecuteAsync();

                return result.AccessToken;
            }

            try
            {
                var result = await publicClientApplication
                    .AcquireTokenSilent(new[] { "Mail.Send" }, account)
                    .ExecuteAsync();

                return result.AccessToken;
            }
            catch (MsalUiRequiredException)
            {
                var result = await publicClientApplication
                    .AcquireTokenInteractive(new[] { "Mail.Send" })
                    .ExecuteAsync();

                return result.AccessToken;
            }
        }

        public async System.Threading.Tasks.Task SendEmailAsync(TextBox textBoxAsu)
        {
            var accessToken = await GetAccessTokenAsync();
            string nomeUsuario = Environment.UserName;
            nomeUsuario = nomeUsuario.Replace('.', ' ');
            nomeUsuario = string.Join(" ", nomeUsuario.Split(' ').Select(p => char.ToUpper(p[0]) + p.Substring(1).ToLower()));

            var subject = textBoxAsu.Text;
            string saudacao = GetSaudacao();
            string corpoEmail = "<html><body contenteditable=\"false\">";
            corpoEmail += "<p style=\"font-family: Calibri; font-size: 14px;\">" + CorpoEmailImagens + "</p>";
            corpoEmail += "</body></html>";
            var attachments = new List<Dictionary<string, object>>();

            foreach (string nomeArquivo in listBoxFicheiros.Items)
            {
                if (ficheirosSelecionados.TryGetValue(nomeArquivo.ToString(), out string caminhoCompleto))
                {
                    byte[] fileBytes = File.ReadAllBytes(caminhoCompleto);
                    string contentBytes = Convert.ToBase64String(fileBytes);
                    string contentType = GetMimeType(Path.GetExtension(caminhoCompleto));

                    if (string.IsNullOrEmpty(contentType))
                        contentType = "application/octet-stream";

                    var attachment = new Dictionary<string, object>
                    {
                        ["@odata.type"] = "#microsoft.graph.fileAttachment",
                        ["name"] = nomeArquivo,
                        ["contentBytes"] = contentBytes,
                        ["contentType"] = contentType
                    };

                    attachments.Add(attachment);
                }
            }

            var destinatarios = listBoxPara.Items
                .Cast<string>()
                .Select(item => item.Trim())
                .Where(item => !string.IsNullOrEmpty(item))
                .Where(item => IsValidEmail(item))
                .Distinct()
                .ToList();

            var toRecipients = destinatarios
                .Select(email => new { EmailAddress = new { Address = email } })
                .ToArray();

            var ccDestinatarios = listBoxCC.Items
                .Cast<string>()
                .Select(email => email.Trim())
                .Where(email => !string.IsNullOrEmpty(email))
                .Where(email => IsValidEmail(email))
                .Distinct()
                .ToList();

            var ccRecipientsFormatted = ccDestinatarios
                .Select(email => new { EmailAddress = new { Address = email } })
                .ToArray();

            var emailMessage = new
            {
                Message = new
                {
                    Subject = subject,
                    Body = new
                    {
                        ContentType = "HTML",
                        Content = corpoEmail
                    },
                    ToRecipients = toRecipients,
                    CcRecipients = ccRecipientsFormatted,
                    Attachments = attachments
                },
                SaveToSentItems = true
            };

            var jsonContent = Newtonsoft.Json.JsonConvert.SerializeObject(emailMessage);
            var content = new StringContent(jsonContent, Encoding.UTF8, "application/json");

            using (var client = new HttpClient())
            {
                client.DefaultRequestHeaders.Add("Authorization", $"Bearer {accessToken}");
                client.DefaultRequestHeaders.Add("Accept", "application/json");

                var response = await client.PostAsync("https://graph.microsoft.com/v1.0/me/sendMail", content);

                if (response.IsSuccessStatusCode)
                {
                    CustomGradientPanel4.Visible = true;

                    labelEmailInformação.ForeColor = Color.White;
                    labelEmailInformação.Location = new Point(60, 8);

                    labelEmailInformação.Text = "E-mail enviado com sucesso!";
                    await System.Threading.Tasks.Task.Delay(3000);
                    this.Close();
                }
                else
                {
                    var errorResponse = await response.Content.ReadAsStringAsync();
                    CustomGradientPanel4.Visible = true;
                    CustomGradientPanel4.FillColor = Color.Maroon;
                    CustomGradientPanel4.FillColor2 = Color.Maroon;
                    CustomGradientPanel4.FillColor3 = Color.Maroon;
                    CustomGradientPanel4.FillColor4 = Color.Maroon;
                    labelEmailInformação.Text = "Erro ao enviar o e-mail!";
                    MessageBox.Show(this, $"Erro ao enviar o e-mail: {response.StatusCode} - {errorResponse}");
                }
            }
        }

        private bool IsValidEmail(string email)
        {
            try
            {
                var mailAddress = new System.Net.Mail.MailAddress(email);
                return mailAddress.Address == email;
            }
            catch
            {
                return false;
            }
        }

        private void guna2Button1_Click(object sender, EventArgs e)
        {
            Abrirpasta();
        }

        private void Abrirpasta()
        {
            string numeroobra = label11.Text;
            string ano = "20" + numeroobra.Substring(0, 2);
            string initialDirectory = $@"\\marconi\OFELIZ\OFELIZ\OFM\2.AN\2.CM\DP\1 Obras\{ano}\{numeroobra}\1.8 Projeto\1.8.1 Projeto de execução\1.8.1.2 Peças escritas\1.8.1.2.1 Em vigor";

            if (!Directory.Exists(initialDirectory))
            {
                MessageBox.Show(this, $"Diretório não encontrado:\n{initialDirectory}", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            OpenFileDialog openFileDialog = new OpenFileDialog
            {
                Title = "Selecionar Arquivo(s)",
                Filter = "Todos os Arquivos Suportados|" +
                         "*.pdf;*.xlsm;*.xlsx;*.xls;*.doc;*.docx;*.ppt;*.pptx;*.dwg;*.txt;*.csv;*.jpg;*.jpeg;*.png;*.gif;*.bmp;*.ifc|" +
                         "Documentos PDF|*.pdf|" +
                         "Planilhas Excel|*.xlsm;*.xlsx;*.xls|" +
                         "Documentos Word|*.doc;*.docx|" +
                         "Apresentações PowerPoint|*.ppt;*.pptx|" +
                         "Imagens|*.jpg;*.jpeg;*.png;*.gif;*.bmp|" +
                         "Desenhos AutoCAD (*.dwg)|*.dwg|" +
                         "Textos|*.txt;*.csv|" +
                         "Modelos IFC|*.ifc|" +
                         "Todos os Arquivos|*.*",
                InitialDirectory = initialDirectory,
                Multiselect = true
            };

            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                string[] arquivosSelecionados = openFileDialog.FileNames;

                foreach (string caminhoArquivo in arquivosSelecionados)
                {
                    string nomeArquivo = Path.GetFileName(caminhoArquivo);

                    if (!listBoxFicheiros.Items.Contains(nomeArquivo))
                    {
                        listBoxFicheiros.Items.Add(nomeArquivo);
                        ficheirosSelecionados[nomeArquivo] = caminhoArquivo;
                    }
                }
            }
        }

        private void Frm_Corpo_de_Texto_Email_Pesclarecimento_DragDrop(object sender, DragEventArgs e)
        {
            string[] arquivosArrastados = (string[])e.Data.GetData(DataFormats.FileDrop);

            foreach (string caminhoArquivo in arquivosArrastados)
            {
                string extensao = Path.GetExtension(caminhoArquivo).ToLower();

                if (extensao == ".pdf" || extensao == ".xlsm" || extensao == ".xlsx" || extensao == ".xls" ||
                    extensao == ".doc" || extensao == ".docx" || extensao == ".ppt" || extensao == ".pptx" ||
                    extensao == ".dwg" || extensao == ".txt" || extensao == ".csv" ||
                    extensao == ".jpg" || extensao == ".jpeg" || extensao == ".png" || extensao == ".gif" || extensao == ".bmp")
                {
                    string nomeArquivo = Path.GetFileName(caminhoArquivo);

                    if (!listBoxFicheiros.Items.Contains(nomeArquivo))
                    {
                        listBoxFicheiros.Items.Add(nomeArquivo);
                        ficheirosSelecionados[nomeArquivo] = caminhoArquivo;
                    }
                }
                else
                {
                    MessageBox.Show("Tipo de arquivo não suportado: " + extensao, "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void Frm_Corpo_de_Texto_Email_Pesclarecimento_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                e.Effect = DragDropEffects.Copy;
            }
            else
            {
                e.Effect = DragDropEffects.None;
            }
        }


       
    }

}