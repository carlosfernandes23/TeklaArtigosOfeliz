using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Tekla.Structures.Drawing;
using Tekla.Structures.Filtering;
using Tekla.Structures.Filtering.Categories;
using Tekla.Structures.Geometry3d;
using Tekla.Structures.Model;
using TSM = Tekla.Structures.Model;

namespace TeklaArtigosOfeliz
{
    public partial class Frm_Parafusaria_v2: Form
    {
        public Frm_Parafusaria_v2()
        {
            InitializeComponent();
        }

        private void Frm_Parafusaria_v2_Load(object sender, EventArgs e)
        {
            CarregarDados();
        }

        private void CarregarDados()
        {
            TSM.Model modelo = new TSM.Model();
            string NumeroObra = modelo.GetProjectInfo().ProjectNumber;
            string Fabricante = modelo.GetProjectInfo().Builder;
            string Dataatual = DateTime.Now.ToString("dd/MM/yyyy");
            string NomeModelo = modelo.GetInfo().ModelName;

            string classeEX = string.Empty;
            string lote = string.Empty;
            string dataObra = string.Empty;
            string fase = string.Empty;

            try
            {
                ArrayList peças = new ArrayList(ComunicaTekla.ListadePecasdoConjSelec());

                if (peças.Count == 0)
                {
                    MessageBox.Show(this, "Selecione as peças no Tekla em modo conjunto.");
                    return;
                }

                foreach (TSM.Part peca in peças)
                {
                    if (peca != null)
                    {
                        peca.GetReportProperty("PROJECT.USERDEFINED.PROJECT_USERFIELD_2", ref classeEX);
                        peca.GetReportProperty("USERDEFINED.lote_number", ref lote);
                        peca.GetReportProperty("USERDEFINED.lote_data", ref dataObra);
                        peca.GetReportProperty("USERDEFINED.Fase", ref fase);
                        break; 
                    }
                }

                dataGridView1.Rows.Clear();

                Dictionary<string, int> boltAgrupados = new Dictionary<string, int>();
                Dictionary<string, int> washerAgrupados = new Dictionary<string, int>();
                Dictionary<string, int> nutAgrupados = new Dictionary<string, int>();

                TSM.ModelObjectEnumerator selecionados = new TSM.UI.ModelObjectSelector().GetSelectedObjects();

                while (selecionados.MoveNext())
                {
                    var obj = selecionados.Current;

                    if (obj is TSM.BoltGroup bolt)
                    {
                        string boltStandard = "";
                        string boltSize = "";
                        string boltLength = "";
                        string washerStandard = "";
                        string nutStandard = "";

                        bolt.GetReportProperty("BOLT_STANDARD", ref boltStandard);
                        bolt.GetReportProperty("BOLT_SIZE", ref boltSize);
                        bolt.GetReportProperty("BOLT_LENGTH", ref boltLength);
                        bolt.GetReportProperty("WASHER_TYPE", ref washerStandard);  // ou "WASHER_STANDARD"
                        bolt.GetReportProperty("NUT_TYPE", ref nutStandard);        // ou "NUT_STANDARD"

                        int quantidade = bolt.BoltPositions.Count;

                        // Parafuso
                        string chaveBolt = $"{boltStandard} - {boltSize} x {boltLength}";
                        if (boltAgrupados.ContainsKey(chaveBolt))
                            boltAgrupados[chaveBolt] += quantidade;
                        else
                            boltAgrupados[chaveBolt] = quantidade;

                        // Anilha
                        if (!string.IsNullOrEmpty(washerStandard))
                        {
                            string chaveWasher = $"{washerStandard} - {boltSize}";
                            if (washerAgrupados.ContainsKey(chaveWasher))
                                washerAgrupados[chaveWasher] += quantidade;
                            else
                                washerAgrupados[chaveWasher] = quantidade;
                        }

                        // Porca
                        if (!string.IsNullOrEmpty(nutStandard))
                        {
                            string chaveNut = $"{nutStandard} - {boltSize}";
                            if (nutAgrupados.ContainsKey(chaveNut))
                                nutAgrupados[chaveNut] += quantidade;
                            else
                                nutAgrupados[chaveNut] = quantidade;
                        }
                    }
                }

                foreach (var grupo in boltAgrupados)
                {
                    int rowIndex = dataGridView1.Rows.Add();
                    var row = dataGridView1.Rows[rowIndex];
                    row.Cells["Perfil"].Value = grupo.Key;
                    row.Cells["Qt"].Value = grupo.Value;
                }

                foreach (var grupo in washerAgrupados)
                {
                    int rowIndex = dataGridView1.Rows.Add();
                    var row = dataGridView1.Rows[rowIndex];
                    row.Cells["Perfil"].Value = grupo.Key;
                    row.Cells["Qt"].Value = grupo.Value;
                }

                foreach (var grupo in nutAgrupados)
                {
                    int rowIndex = dataGridView1.Rows.Add();
                    var row = dataGridView1.Rows[rowIndex];
                    row.Cells["Perfil"].Value = grupo.Key;
                    row.Cells["Qt"].Value = grupo.Value;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, "Erro ao carregar dados: " + ex.Message);
                return;
            }

            labelbuilder.Text = Fabricante;
            labelname.Text = NomeModelo;
            labelNobra.Text = NumeroObra;
            labelData.Text = Dataatual;
            labelClasseEx.Text = classeEX;

            AtribuirValores(fase, lote, dataObra);
        }     

        private void AtribuirValores(string fase, string lote , string dataObra)
        {
            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                if (!row.IsNewRow)
                {
                    row.Cells["Fase"].Value = fase;
                    row.Cells["lote"].Value = lote;
                    row.Cells["Entrega"].Value = dataObra;
                }
            }
        }

    }

}

