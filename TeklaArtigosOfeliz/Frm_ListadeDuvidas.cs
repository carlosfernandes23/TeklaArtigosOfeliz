using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net.Http.Headers;
using System.Net.Http;
using System.Text;
using System.Text.Json;
using System.Threading.Tasks;
using System.Windows.Forms;
using ClosedXML.Excel;
using TSM = Tekla.Structures.Model;
using System.Net.Http.Headers;
using System.Text.Json;
using Microsoft.Identity.Client;
using HtmlAgilityPack;
using HtmlDocument = HtmlAgilityPack.HtmlDocument;
using System.Net;
using System.Text.RegularExpressions;
using Org.BouncyCastle.Ocsp;
using System.Threading;

namespace TeklaArtigosOfeliz
{
    public partial class Frm_ListadeDuvidas: Form
    {
        public Frm_ListadeDuvidas()
        {
            InitializeComponent();
                    
        }

        private void Frm_ListadeDuvidas_Load(object sender, EventArgs e)
        {
            TSM.Model modelo = new TSM.Model();
            string numeroobra = modelo.GetProjectInfo().ProjectNumber;
            string nomeobra = modelo.GetProjectInfo().Name;
            string fabricante = modelo.GetProjectInfo().Builder;
            lbl_numeroobra.Text = numeroobra;
            lbl_nomeobra.Text = nomeobra;
            lbl_nomefabricante.Text = fabricante;
            VerificarOuCriarTabela(numeroobra);
            CriarPastas();
            CarregarDadosEImagens();
            AtualizarLblPEsclarecimentoComUltimoPE();
            PreencherComboBoxes();
            PreencherListBoxPE();
        }

        private void guna2Button1_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(txt_duvida.Text))
            {
                MessageBox.Show("O campo de dúvida não pode estar vazio.", "Atenção", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            bool resultado = CriarDuvida();
            CriarPastasParaDuvida();
            CarregarDadosEImagens();
            AtualizarLblPEsclarecimentoComUltimoPE();
            PreencherListBoxPE();
            PreencherComboBoxes();

            if (resultado)
            {
                MessageBox.Show("Dúvida inserida com sucesso!");
            }
            else
            {
                MessageBox.Show("Erro ao inserir a dúvida.");
            }
        }

        private void guna2ImageButton3_Click(object sender, EventArgs e)
        {
            AnexarFicheiros();
        }

        private void guna2ImageButton2_Click(object sender, EventArgs e)
        {
            CapturarEGuardarImagem();
        }

        private void buttonLimaprlistboxprint_Click(object sender, EventArgs e)
        {
            listBoxPrint.Items.Clear();

        }

        private void guna2Button2_Click(object sender, EventArgs e)
        {
            FiltrarDataGridView();
        }

        private void guna2Button4_Click(object sender, EventArgs e)
        {
            GerarEmailComDuvidasSelecionadas();
        }

        private async void guna2Button3_Click(object sender, EventArgs e)
        {
            string datahoje = DateTime.Now.ToString("dd/MM/yyyy");
            string texto = richTextBoxRespostas.Text;

            List<string> listaRespostas = ExtrairRespostasFormatadas(texto);

            foreach (var resposta in listaRespostas)
            {
                var match = Regex.Match(resposta, @"^(?<num>\d{3})\.\s*(?<conteudo>.+)", RegexOptions.Singleline);
                if (match.Success)
                {
                    string numPE = match.Groups["num"].Value;
                    string respostaLimpa = match.Groups["conteudo"].Value;
                    GerarResposta(respostaLimpa, numPE, datahoje);
                }
            }
            CarregarDadosEImagens();
        }

        private void PreencherListBoxPE()
        {
            listBoxPE.Items.Clear(); 

            foreach (DataGridViewRow row in dataGridViewBD.Rows)
            {
                if (!row.IsNewRow)
                {
                    var valor = row.Cells[0].Value?.ToString();

                    if (!string.IsNullOrWhiteSpace(valor))
                    {
                        listBoxPE.Items.Add(valor);
                    }
                }
            }
        }

        private void AtualizarLblPEsclarecimentoComUltimoPE()
        {
            if (dataGridViewBD.Rows.Count == 0)
            {
                lbl_pesclarecimento.Text = "001";
                return;
            }

            for (int i = dataGridViewBD.Rows.Count - 1; i >= 0; i--)
            {
                var row = dataGridViewBD.Rows[i];

                if (!row.IsNewRow)
                {
                    var valorPE = row.Cells["NumPE"].Value?.ToString();

                    if (!string.IsNullOrWhiteSpace(valorPE) && int.TryParse(valorPE, out int numeroPE))
                    {
                        int novoPE = numeroPE + 1;
                        lbl_pesclarecimento.Text = novoPE.ToString("D3"); 
                        return;
                    }
                }
            }

            lbl_pesclarecimento.Text = "CUIDADO !!!PE não encontrado";
        }

        private string CriarPastas()
        {
            string numeroObra = lbl_numeroobra.Text;
            string ano = "20" + numeroObra.Substring(0, 2);

            string caminhoBase = Path.Combine(
                @"\\marconi\OFELIZ\OFELIZ\OFM\2.AN\2.CM\DP\1 Obras",
                ano,
                numeroObra,
                @"1.8 Projeto\1.8.1 Projeto de execução\1.8.1.2 Peças escritas\1.8.1.2.1 Em vigor\Duvidas"
            );

            if (!Directory.Exists(caminhoBase))
            {
                Directory.CreateDirectory(caminhoBase);
            }

            return caminhoBase;
        }

        private string CriarPastasParaDuvida()
        {
            string numeroObra = lbl_numeroobra.Text;
            string numeroDuvida = lbl_pesclarecimento.Text;
            string ano = "20" + numeroObra.Substring(0, 2);

            string caminhoBase = Path.Combine(
                @"\\marconi\OFELIZ\OFELIZ\OFM\2.AN\2.CM\DP\1 Obras",
                ano,
                numeroObra,
                @"1.8 Projeto\1.8.1 Projeto de execução\1.8.1.2 Peças escritas\1.8.1.2.1 Em vigor\Duvidas",
                numeroDuvida
            );

            if (!Directory.Exists(caminhoBase))
            {
                Directory.CreateDirectory(caminhoBase);
            }

            string caminhoBaseEnviado = Path.Combine(caminhoBase, "Enviado");
            if (!Directory.Exists(caminhoBaseEnviado))
            {
                Directory.CreateDirectory(caminhoBaseEnviado);
            }

            string caminhoBaseResposta = Path.Combine(caminhoBase, "Resposta");
            if (!Directory.Exists(caminhoBaseResposta))
            {
                Directory.CreateDirectory(caminhoBaseResposta);
            }

            return caminhoBase;
        }

        private void CarregarDadosEImagens()
        {
            var dt = BuscarDadosDaBase();
            ConfigurarDataGridView(dt);
            PreencherAnexosEnvio();
            PreencherAnexosResposta();
        }

        private DataTable BuscarDadosDaBase()
        {
            string numeroObra = lbl_numeroobra.Text.Trim();
            string nometable = "dbo.N_" + numeroObra;

            ComunicaBaseDadosListaDuvidas BD = new ComunicaBaseDadosListaDuvidas();
            try
            {
                BD.ConectarBDduvida();
                string query = $"SELECT ID, NumPE, Preparador, Duvida, DataEnvio, Resposta, DataResposta, Conclusao FROM {nometable}";
                DataTable dt = BD.Procurarbdduvida(query);
                return dt;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Erro ao conectar à base de dados: " + ex.Message);
                return null;
            }
            finally
            {
                BD.DesonectarBDduvida();
            }
        }

        private void ConfigurarDataGridView(DataTable dataTable)
        {
            if (dataTable == null) return;

            dataGridViewBD.Columns.Clear();
            dataGridViewBD.AutoGenerateColumns = false;
            dataGridViewBD.AllowUserToAddRows = false;
            dataGridViewBD.AllowUserToDeleteRows = false;
            dataGridViewBD.ReadOnly = true;

            dataGridViewBD.Columns.Add(new DataGridViewTextBoxColumn() { Name = "NumPE", HeaderText = "Nº PE", DataPropertyName = "NumPE", Width = 60 });
            dataGridViewBD.Columns.Add(new DataGridViewTextBoxColumn() { Name = "Preparador", HeaderText = "Preparador", DataPropertyName = "Preparador", Width = 120 });
            dataGridViewBD.Columns.Add(new DataGridViewTextBoxColumn() { Name = "Duvida", HeaderText = "Dúvida", DataPropertyName = "Duvida", Width = 700 });

            dataGridViewBD.Columns.Add(new DataGridViewImageColumn() { Name = "AnexoImagem", HeaderText = "Anexo / Imagem Envio", Width = 650, ImageLayout = DataGridViewImageCellLayout.Zoom });
            dataGridViewBD.Columns.Add(new DataGridViewTextBoxColumn() { Name = "DataEnvio", HeaderText = "Data de Envio", DataPropertyName = "DataEnvio", Width = 100 });
            dataGridViewBD.Columns.Add(new DataGridViewTextBoxColumn() { Name = "Resposta", HeaderText = "Resposta", DataPropertyName = "Resposta", Width = 700 });

            dataGridViewBD.Columns.Add(new DataGridViewImageColumn() { Name = "AnexoImagem2", HeaderText = "Anexo / Imagem Resposta", Width = 650, ImageLayout = DataGridViewImageCellLayout.Zoom });
            dataGridViewBD.Columns.Add(new DataGridViewTextBoxColumn() { Name = "DataResposta", HeaderText = "Data da Resposta", DataPropertyName = "DataResposta", Width = 100 });
            dataGridViewBD.Columns.Add(new DataGridViewTextBoxColumn() { Name = "Conclusao", HeaderText = "Conclusão", DataPropertyName = "Conclusao", Width = 150 });

            dataGridViewBD.DataSource = dataTable;

            if (dataGridViewBD.Columns.Contains("ID"))
            {
                dataGridViewBD.Columns["ID"].Visible = false;
            }
            foreach (DataGridViewColumn col in dataGridViewBD.Columns)
            {
                col.SortMode = DataGridViewColumnSortMode.NotSortable;
            }
            foreach (DataGridViewRow row in dataGridViewBD.Rows)
            {
                row.Height = 250;
            }
        }

        private Image CriarImagemComposta(string caminhoPasta)
        {
            if (!Directory.Exists(caminhoPasta))
                return null;

            var arquivos = Directory.GetFiles(caminhoPasta)
                .Where(f => f.EndsWith(".jpg", StringComparison.OrdinalIgnoreCase) ||
                            f.EndsWith(".png", StringComparison.OrdinalIgnoreCase) ||
                            f.EndsWith(".jpeg", StringComparison.OrdinalIgnoreCase))
                .Take(5)  
                .ToList();

            if (arquivos.Count == 0)
                return null;

            int thumbWidth = 500;
            int thumbHeight = 300;
            int spacing = 5;
            int larguraTotal = (thumbWidth + spacing) * arquivos.Count - spacing;
            int alturaTotal = thumbHeight;

            Bitmap imagemComposta = new Bitmap(larguraTotal, alturaTotal);
            using (Graphics g = Graphics.FromImage(imagemComposta))
            {
                g.Clear(Color.White);

                for (int i = 0; i < arquivos.Count; i++)
                {
                    try
                    {
                        using (var stream = new FileStream(arquivos[i], FileMode.Open, FileAccess.Read))
                        using (var imgOriginal = Image.FromStream(stream))
                        {
                            var thumb = new Bitmap(thumbWidth, thumbHeight);
                            using (Graphics tg = Graphics.FromImage(thumb))
                            {
                                tg.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
                                tg.DrawImage(imgOriginal, 0, 0, thumbWidth, thumbHeight);
                            }
                            g.DrawImage(thumb, i * (thumbWidth + spacing), 0);
                            thumb.Dispose();
                        }
                    }
                    catch
                    {
                        // Ignorar imagens com erro
                    }
                }
            }

            return imagemComposta;
        }

        private void PreencherAnexosEnvio()
        {
            string numeroObra = lbl_numeroobra.Text.Trim();
            string ano = "20" + numeroObra.Substring(0, 2);

            foreach (DataGridViewRow row in dataGridViewBD.Rows)
            {
                if (row.IsNewRow) continue;

                string numPE = row.Cells["NumPE"].Value?.ToString().PadLeft(3, '0');
                if (string.IsNullOrEmpty(numPE)) continue;

                string caminhoAnexoEnvio = $@"\\marconi\OFELIZ\OFELIZ\OFM\2.AN\2.CM\DP\1 Obras\{ano}\{numeroObra}\1.8 Projeto\1.8.1 Projeto de execução\1.8.1.2 Peças escritas\1.8.1.2.1 Em vigor\Duvidas\{numPE}\Enviado";

                Image imgComposta = CriarImagemComposta(caminhoAnexoEnvio);    
                if (imgComposta != null && row.Cells["AnexoImagem"] is DataGridViewImageCell cell)
                {
                    cell.Value = imgComposta;
                }
                else if (row.Cells["AnexoImagem"] is DataGridViewImageCell cellVazia)
                {
                    cellVazia.Value = null;
                    cellVazia.Style.NullValue = null; 
                }

            }
        }

        private void PreencherAnexosResposta()
        {
            string numeroObra = lbl_numeroobra.Text.Trim();
            string ano = "20" + numeroObra.Substring(0, 2);

            foreach (DataGridViewRow row in dataGridViewBD.Rows)
            {
                if (row.IsNewRow) continue;

                string numPE = row.Cells["NumPE"].Value?.ToString().PadLeft(3, '0');
                if (string.IsNullOrEmpty(numPE)) continue;

                string caminhoAnexoResposta = $@"\\marconi\OFELIZ\OFELIZ\OFM\2.AN\2.CM\DP\1 Obras\{ano}\{numeroObra}\1.8 Projeto\1.8.1 Projeto de execução\1.8.1.2 Peças escritas\1.8.1.2.1 Em vigor\Duvidas\{numPE}\Resposta";

                Image imgComposta = CriarImagemComposta(caminhoAnexoResposta);
                if (imgComposta != null && row.Cells["AnexoImagem2"] is DataGridViewImageCell cell)
                {
                    cell.Value = imgComposta;
                }
                else if (row.Cells["AnexoImagem2"] is DataGridViewImageCell cellVazia)
                {
                    cellVazia.Value = null;
                    cellVazia.Style.NullValue = null; 
                }
            }
        }

        private void CapturarEGuardarImagem()
        {
            CriarPastasParaDuvida();
            string numeroduvida = lbl_pesclarecimento.Text;
            string numeroObra = lbl_numeroobra.Text;
            string ano = "20" + numeroObra.Substring(0, 2);

            string caminhoDestino = $@"\\marconi\OFELIZ\OFELIZ\OFM\2.AN\2.CM\DP\1 Obras\{ano}\{numeroObra}\1.8 Projeto\1.8.1 Projeto de execução\1.8.1.2 Peças escritas\1.8.1.2.1 Em vigor\Duvidas\{numeroduvida}\Enviado";

            Process.Start("explorer.exe", "ms-screenclip:");

            Image img = null;
            int tempoMaximoMs = 30000;
            int intervaloVerificacaoMs = 500;
            int tempoPassado = 0;

            while (tempoPassado < tempoMaximoMs)
            {
                System.Threading.Thread.Sleep(intervaloVerificacaoMs);
                tempoPassado += intervaloVerificacaoMs;

                if (Clipboard.ContainsImage())
                {
                    img = Clipboard.GetImage();
                    break;
                }
            }

            if (img != null)
            {
                var existentes = Directory.GetFiles(caminhoDestino, $"{numeroduvida}_*.jpg");
                int proximoNumero = existentes.Length + 1;

                string nomeArquivo = $"{numeroduvida}_{proximoNumero}.jpg";
                string caminhoCompleto = Path.Combine(caminhoDestino, nomeArquivo);

                img.Save(caminhoCompleto, System.Drawing.Imaging.ImageFormat.Jpeg);
                listBoxPrint.Items.Add(nomeArquivo);
                MessageBox.Show($"Imagem salva como: {nomeArquivo}", "Sucesso", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                MessageBox.Show("Nenhuma imagem foi capturada no tempo limite!", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private bool CriarDuvida()
        {
            string textoDuvida = txt_duvida.Text;
            string numeroObra = lbl_numeroobra.Text;
            string numeropedido = lbl_pesclarecimento.Text;
            string user = Environment.UserName;
            string userFormatado = string.Join(" ",
                user
                    .Replace('.', ' ')
                    .Split(' ')
                    .Where(p => !string.IsNullOrWhiteSpace(p))
                    .Select(p => char.ToUpper(p[0]) + p.Substring(1).ToLower())
            );
            string conclusao = "Aguardar Resposta";
            string datahoje = DateTime.Now.ToString("dd/MM/yyyy");
            string nometable = "dbo.N_" + numeroObra;
            string query = $"INSERT INTO {nometable} (NumPE, Preparador, Duvida, DataEnvio, Conclusao) " +
                                             "VALUES (@NumPE, @Preparador, @Duvida, @DataEnvio, @Conclusao)";

            ComunicaBaseDadosListaDuvidas BD = new ComunicaBaseDadosListaDuvidas();

            try
            {
                BD.ConectarBDduvida();
                using (SqlCommand cmd = new SqlCommand(query, BD.GetConnection()))
                {
                    cmd.Parameters.AddWithValue("@NumPE", numeropedido);
                    cmd.Parameters.AddWithValue("@Preparador", userFormatado);
                    cmd.Parameters.AddWithValue("@Duvida", textoDuvida);
                    cmd.Parameters.AddWithValue("@DataEnvio", datahoje);
                    cmd.Parameters.AddWithValue("@Conclusao", conclusao);
                    

                    cmd.ExecuteNonQuery();
                }
                return true;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Erro ao registrar a hora de início: " + ex.Message);
                return false;
            }
            finally
            {
                BD.DesonectarBDduvida();
            }
        }

        private void VerificarOuCriarTabela(string numeroobra)
        {
            string nomeTabela = $"N_{numeroobra}";
            string queryVerificar = $"SELECT COUNT(*) FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME = '{nomeTabela}'";

            string queryCriar = $@"
                                CREATE TABLE {nomeTabela} (
                                    ID INT IDENTITY(1,1) PRIMARY KEY,
                                    NumPE NVARCHAR(MAX) NOT NULL,
                                    Preparador NVARCHAR(MAX) NOT NULL,
                                    Duvida NVARCHAR(MAX) NOT NULL,
                                    DataEnvio NVARCHAR(MAX) NOT NULL,
                                    Resposta NVARCHAR(MAX) NULL,
                                    DataResposta NVARCHAR(MAX) NULL,
                                    Conclusao NVARCHAR(MAX) NULL
                                );";

            ComunicaBaseDadosListaDuvidas BD = new ComunicaBaseDadosListaDuvidas();

            try
            {
                BD.ConectarBDduvida();
                using (SqlCommand cmdVerificar = new SqlCommand(queryVerificar, BD.GetConnection()))
                {
                    int count = (int)cmdVerificar.ExecuteScalar();

                    if (count == 0) 
                    {
                        using (SqlCommand cmdCriar = new SqlCommand(queryCriar, BD.GetConnection()))
                        {
                            cmdCriar.ExecuteNonQuery();
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Erro ao verificar/criar tabela: " + ex.Message);
            }
            finally
            {
                BD.DesonectarBDduvida();
            }
        }

        private void dataGridViewBD_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < 0 || e.ColumnIndex < 0) return;

            string coluna = dataGridViewBD.Columns[e.ColumnIndex].Name;

            if (coluna == "AnexoImagem" || coluna == "AnexoImagem2")
            {
                string numeroObra = lbl_numeroobra.Text.Trim();
                if (string.IsNullOrEmpty(numeroObra)) return;

                string ano = "20" + numeroObra.Substring(0, 2);

                string numPE = dataGridViewBD.Rows[e.RowIndex].Cells["NumPE"].Value?.ToString().PadLeft(3, '0');
                if (string.IsNullOrEmpty(numPE)) return;

                string tipo = (coluna == "AnexoImagem") ? "Enviado" : "Resposta";

                string caminho = $@"\\marconi\OFELIZ\OFELIZ\OFM\2.AN\2.CM\DP\1 Obras\{ano}\{numeroObra}\1.8 Projeto\1.8.1 Projeto de execução\1.8.1.2 Peças escritas\1.8.1.2.1 Em vigor\Duvidas\{numPE}\{tipo}";

                if (Directory.Exists(caminho))
                {
                    Process.Start("explorer.exe", caminho);
                }
                else
                {
                    MessageBox.Show("Pasta não encontrada:\n" + caminho, "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void AnexarFicheiros()
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Title = "Selecionar ficheiros para anexar";
            openFileDialog.Multiselect = true;
            openFileDialog.Filter = "Todos os ficheiros suportados|*.png;*.jpg;*.jpeg;*.pdf;*.ifc;*.docx;*.xlsx;*.dwg;*.txt;*.zip;*.rar|Imagens (*.png, *.jpg, *.jpeg)|*.png;*.jpg;*.jpeg|PDF (*.pdf)|*.pdf|IFC (*.ifc)|*.ifc|Todos os ficheiros (*.*)|*.*";

            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                foreach (string ficheiro in openFileDialog.FileNames)
                {
                    listBoxPrint.Items.Add(ficheiro);
                }
            }       
        }

        private void listBoxPrint_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Delete)
            {
                if (listBoxPrint.SelectedIndex >= 0)
                {
                    listBoxPrint.Items.RemoveAt(listBoxPrint.SelectedIndex);
                }
            }
        }

        private void listBoxPrint_DragEnter(object sender, DragEventArgs e)
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

        private void listBoxPrint_DragDrop(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                string[] ficheiros = (string[])e.Data.GetData(DataFormats.FileDrop);
                foreach (string ficheiro in ficheiros)
                {
                    listBoxPrint.Items.Add(ficheiro);
                }
            }
        }

        private void PreencherComboBoxes()
        {
            comboBoxNpe.Items.Clear();
            comboBoxPrep.Items.Clear();

            HashSet<string> nomesUnicos = new HashSet<string>(); 

            foreach (DataGridViewRow row in dataGridViewBD.Rows)
            {
                if (!row.IsNewRow)
                {
                    var valorColuna0 = row.Cells[0].Value?.ToString();
                    if (!string.IsNullOrEmpty(valorColuna0))
                    {
                        comboBoxNpe.Items.Add(valorColuna0);
                    }

                    var valorColuna1 = row.Cells[1].Value?.ToString();
                    if (!string.IsNullOrEmpty(valorColuna1) && nomesUnicos.Add(valorColuna1))
                    {
                        comboBoxPrep.Items.Add(valorColuna1);
                    }
                }
            }
        }

        private void FiltrarDataGridView()
        {
            string filtroNpe = comboBoxNpe.SelectedItem?.ToString();
            string filtroPrep = comboBoxPrep.SelectedItem?.ToString();
            string filtroConcl = comboBoxConcl.SelectedItem?.ToString();

            CurrencyManager currencyManager = (CurrencyManager)BindingContext[dataGridViewBD.DataSource];
            currencyManager.SuspendBinding(); 

            foreach (DataGridViewRow row in dataGridViewBD.Rows)
            {
                if (row.IsNewRow) continue;

                bool mostrarLinha = true;

                if (!string.IsNullOrEmpty(filtroNpe))
                {
                    string valorColuna0 = row.Cells[0].Value?.ToString();
                    if (valorColuna0 != filtroNpe)
                        mostrarLinha = false;
                }

                if (!string.IsNullOrEmpty(filtroPrep))
                {
                    string valorColuna1 = row.Cells[1].Value?.ToString();
                    if (valorColuna1 != filtroPrep)
                        mostrarLinha = false;
                }

                if (!string.IsNullOrEmpty(filtroConcl))
                {
                    string valorColuna8 = row.Cells[8].Value?.ToString();
                    if (valorColuna8 != filtroConcl)
                        mostrarLinha = false;
                }

                row.Visible = mostrarLinha;
            }

            currencyManager.ResumeBinding(); 
        }
              
        private  bool GerarResposta(string resposta, string numPE, string dataFormatada)
        {
            string numeroobra = lbl_numeroobra.Text.Trim();
            string conclusao = "Concluido";
            string nometable = "dbo.N_" + numeroobra;
                                   
            string query = $"UPDATE {nometable} " +
                           "SET Resposta = @Resposta, DataResposta = @DataResposta, Conclusao = @Conclusao " +
                           "WHERE NumPE = @NumPE";

            ComunicaBaseDadosListaDuvidas BD = new ComunicaBaseDadosListaDuvidas();
            try
            {
                BD.ConectarBDduvida();
                using (SqlCommand cmd = new SqlCommand(query, BD.GetConnection()))
                {
                    cmd.Parameters.AddWithValue("@NumPE", numPE);
                    cmd.Parameters.AddWithValue("@Resposta", resposta);
                    cmd.Parameters.AddWithValue("@DataResposta", dataFormatada);
                    cmd.Parameters.AddWithValue("@Conclusao", conclusao);

                    cmd.ExecuteNonQuery();
                }
                return true;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Erro ao registrar dúvida: " + ex.Message);
                return false;
            }
            finally
            {
                BD.DesonectarBDduvida();
            }
        }

        public static List<string> ExtrairRespostasFormatadas(string texto)
        {
            var respostas = new List<string>();
            var partes = Regex.Split(texto.Trim(), @"\.\s*(?=\n|$)");

            int contador = 1;
            foreach (var parte in partes)
            {
                string respostaLimpa = parte.Trim();
                if (!string.IsNullOrWhiteSpace(respostaLimpa))
                {
                    string numeroFormatado = contador.ToString("D3");
                    respostas.Add($"{numeroFormatado}. {respostaLimpa}.");
                    contador++;
                }
            }

            return respostas;
        }

        private DataTable BuscarDadosDaBasePorNumPE(List<string> numerosPE)
        {
            if (numerosPE == null || numerosPE.Count == 0)
                return null;

            string numeroObra = lbl_numeroobra.Text.Trim();
            string nometable = "dbo.N_" + numeroObra;
            string filtroNumPE = string.Join(",", numerosPE.Select(pe => $"'{pe}'"));

            string query = $"SELECT ID, NumPE, Preparador, Duvida, DataEnvio, Resposta, DataResposta, Conclusao " +
                           $"FROM {nometable} WHERE NumPE IN ({filtroNumPE})";

            ComunicaBaseDadosListaDuvidas BD = new ComunicaBaseDadosListaDuvidas();
            try
            {
                BD.ConectarBDduvida();
                DataTable dt = BD.Procurarbdduvida(query);
                return dt;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Erro ao conectar à base de dados: " + ex.Message);
                return null;
            }
            finally
            {
                BD.DesonectarBDduvida();
            }
        }

        private List<string> ObterNumPESelecionados()
        {
            return listBoxPE.SelectedItems.Cast<string>().ToList();
        }

        private void GerarEmailComDuvidasSelecionadas()
        {
            List<string> numerosSelecionados = ObterNumPESelecionados();
            if (numerosSelecionados.Count == 0)
            {
                MessageBox.Show("Por favor, selecione pelo menos um PE na lista.");
                return;
            }
            string saudacao = GetSaudacao();

            DataTable dados = BuscarDadosDaBasePorNumPE(numerosSelecionados);
            if (dados == null || dados.Rows.Count == 0)
            {
                MessageBox.Show("Nenhum dado encontrado para os PE selecionados.");
                return;
            }
            string numeroObra = lbl_numeroobra.Text;
            string nomeobra = lbl_nomeobra.Text;
            string ano = "20" + numeroObra.Substring(0, 2);

            string subject = numeroObra + " -- " + nomeobra + " --  ESCLARECIMENTO DE DÚVIDAS";
            string imagemOfelizFilePath = @"\\marconi\COMPANY SHARED FOLDER\OFELIZ\OFM\2.AN\2.CM\DP\4 Produção\Desenvolvimentos\Ficheiros Temp tekla artigos (Nao Apagar)\ofeliz_logo.png";
            string nomeUsuario = Environment.UserName;
            nomeUsuario = nomeUsuario.Replace('.', ' ');
            nomeUsuario = string.Join(" ", nomeUsuario.Split(' ').Select(p => char.ToUpper(p[0]) + p.Substring(1).ToLower()));

            string corpoEmail = "<html><body contenteditable=\"false\">";
            corpoEmail += "<p style=\"font-family: Calibri; font-size: 14px;\">" + saudacao + "</p>";
            corpoEmail += "<p style=\"font-family: Calibri; font-size: 14px;\">Venho por este meio solicitar o esclarecimento de dúvidas da obra em assunto. &nbsp;</p>";
            corpoEmail += "<ol style=\"font-family: Calibri; font-size: 14px;\">";

            foreach (DataRow row in dados.Rows)
            {
                string numPE = row["NumPE"].ToString();

                corpoEmail += "<li>";
                corpoEmail += row["Duvida"] + "</li>";

                string caminhoBase = Path.Combine(
                    @"\\marconi\OFELIZ\OFELIZ\OFM\2.AN\2.CM\DP\1 Obras",
                    ano,
                    numeroObra,
                    @"1.8 Projeto\1.8.1 Projeto de execução\1.8.1.2 Peças escritas\1.8.1.2.1 Em vigor\Duvidas",
                    numPE,
                    "Enviado"
                );

                if (Directory.Exists(caminhoBase))
                {
                    string[] imagens = Directory.GetFiles(caminhoBase, "*.*")
                                                .Where(f => f.EndsWith(".png", StringComparison.OrdinalIgnoreCase)
                                                         || f.EndsWith(".jpg", StringComparison.OrdinalIgnoreCase)
                                                         || f.EndsWith(".jpeg", StringComparison.OrdinalIgnoreCase))
                                                .ToArray();

                    foreach (string imagemPath in imagens)
                    {
                        corpoEmail += $"<p><img src='file:///{imagemPath.Replace("\\", "/")}' width='500'></p>";
                    }
                }
                else
                {
                    corpoEmail += "<p style='color: red;'>[Pasta de imagens não encontrada]</p>";
                }
            }

            corpoEmail += "</ol>";
            corpoEmail += "<font face='Calibri' size='3'><p>Melhores Cumprimentos,</p></font>";
            corpoEmail += "<font face = 'Calibri' size = '3' > <b>" + nomeUsuario + "</b> </Font> <br>";
            corpoEmail += "<font face = 'Calibri' size = '3' > Construção Metálica | Preparador </Font> <br>";
            corpoEmail += "<font face = 'Calibri' size = '3' > T + 351 253 080 609 * </font> <br>";
            corpoEmail += "<font color='red' font face = 'Calibri ' size = '3'> ofeliz.com </font> <br>";
            corpoEmail += "<p><a href='https://www.ofeliz.com'><img src='file:///" + imagemOfelizFilePath.Replace("\\", "/") + "' width='127' height='34'></a></p>";
            corpoEmail += "<i><font color='Light grey' font face = 'Calibri ' size = '1.5'> Alvará Nº 10553 – Pub. *Chamada para a rede fixa nacional. </font> </i><br>";
            corpoEmail += "<i><font color='green' font face = 'Calibri ' size = '1.5'> Antes de imprimir este e-mail tenha em consideração o meio ambiente. </font> </i><br>";
            corpoEmail += "</body></html>";

            Frm_Corpo_de_Texto_Email_Pesclarecimento previewForm = new Frm_Corpo_de_Texto_Email_Pesclarecimento("Enviar Dúvidas", corpoEmail, subject);
            previewForm.ShowDialog(this);
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


    }

}

