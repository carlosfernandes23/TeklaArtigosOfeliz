using DocumentFormat.OpenXml.ExtendedProperties;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using TSM = Tekla.Structures.Model;

namespace TeklaArtigosOfeliz
{
    public partial class Frm_AbrirObra : Form
    {
        public Frm_AbrirObra()
        {
            InitializeComponent();
        }

        private void Frm_AbrirObra_Load(object sender, EventArgs e)
        {
            TSM.Model modelo = new TSM.Model();
            string nobra = modelo.GetProjectInfo().ProjectNumber;
            textBoxnobra.Text = nobra;
        }

        private void abrirpastaobra()
        {
            try
            {
                TSM.Model modelo = new TSM.Model();
                string PastaModelo = modelo.GetInfo().ModelPath;

                if (!string.IsNullOrEmpty(PastaModelo) && System.IO.Directory.Exists(PastaModelo))
                {
                    System.Diagnostics.Process.Start("explorer.exe", PastaModelo);
                }
                else
                {
                    MessageBox.Show("A pasta do modelo não foi encontrada.",
                                    "Erro",
                                    MessageBoxButtons.OK,
                                    MessageBoxIcon.Error);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Erro ao abrir a pasta: " + ex.Message,
                                "Erro",
                                MessageBoxButtons.OK,
                                MessageBoxIcon.Error);
            }
        }

        private void abrir(string departamento)
        {
            string inicio = @"\\Marconi\COMPANY SHARED FOLDER\OFELIZ\OFM\2.AN\2.CM\DP\11 Partilhada\";
            string nobra = textBoxnobra.Text.Trim();
            string ano = string.Empty;            
            if (!string.IsNullOrEmpty(nobra) && nobra.Length >= 2)
            {
                string prefixo = nobra.Substring(0, 2); 
                ano = "20" + prefixo; 
            }          

            string caminhoCompleto = System.IO.Path.Combine(inicio, ano, departamento, nobra);

            if (System.IO.Directory.Exists(caminhoCompleto))
            {
                System.Diagnostics.Process.Start("explorer.exe", caminhoCompleto);
            }
            else
            {
                string caminhoAlternativo = System.IO.Path.Combine(inicio, ano, departamento);

                if (System.IO.Directory.Exists(caminhoAlternativo))
                {
                    MessageBox.Show(
                        "A pasta da obra não existe.\nSerá aberto o departamento correspondente.",
                        "Aviso",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning);

                    System.Diagnostics.Process.Start("explorer.exe", caminhoAlternativo);
                }
                else
                {
                    MessageBox.Show(
                        "Nem a pasta da obra nem a pasta do departamento foram encontradas.",
                        "Erro",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Error);
                }
            }
        }

        private void buttonpastaobra_Click(object sender, EventArgs e)
        {
            abrirpastaobra();
        }

        private void Buttonlaser_Click(object sender, EventArgs e)
        {
            string departamento = "LASER";
            abrir(departamento); 
        }

        private void Buttoncq_Click(object sender, EventArgs e)
        {
            string departamento = "CQ";
            abrir(departamento);
        }

        private void Buttoncp_Click(object sender, EventArgs e)
        {
            string departamento = "CP";
            abrir(departamento);
        }

        private void Buttondap_Click(object sender, EventArgs e)
        {
            string departamento = "DAP";
            abrir(departamento);
        }

        private void Buttonarmazem_Click(object sender, EventArgs e)
        {
            string departamento = "ARM";
            abrir(departamento);
        }
    }
}
