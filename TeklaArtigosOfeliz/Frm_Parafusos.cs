using System;
using System.Collections.Generic;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Windows.Forms;
using Microsoft.Office.Interop.Outlook;
using Exception = System.Exception;
using Microsoft.Office.Interop.Excel;
using System.Collections;
using System.Reflection;
using TSM = Tekla.Structures.Model;
using Tekla.Structures.InpParser;
using System.Text.RegularExpressions;

namespace TeklaArtigosOfeliz
{
    public partial class Frm_Parafusos : Form
    {
        Frm_Inico formpai;
        public Frm_Parafusos(Frm_Inico _formpai)
        {
            InitializeComponent();
            formpai = _formpai;
        }

        private void parafusos_Load(object sender, EventArgs e)
        {
            carregadadosTabela1();
            carregadadosTabela2();
            //VerificarColunaClasse();
        }
        public void carregadadosTabela1()
        {
            dataGridView1.Rows.Clear();
            string line = null;
            int i = 1;
            StreamReader file = new StreamReader(@"c:\r\OFELIZ.CSV", Encoding.Default, true);
            while ((line = file.ReadLine()) != null)
            {
                if (i == 2)
                {
                    var fields = line.Split(';');
                    label1.Text = fields[1];
                }
                if (i == 3)
                {
                    var fields = line.Split(';');
                    label2.Text = fields[1];
                }
                if (i == 4)
                {
                    var fields = line.Split(';');
                    lbl_numeroobra.Text = fields[1];
                }
                if (i == 5)
                {
                    var fields = line.Split(';');
                    label4.Text = fields[1];
                }
                if (i == 6)
                {
                    var fields = line.Split(';');
                    label5.Text = fields[1];
                }

                if (i > 8)
                {
                    var fields = line.Split(';');
                    dataGridView1.Rows.Add(fields);
                }
                i++;
            }
            file.Close();
            File.Delete(@"C:\R\OFELIZ.CSV");
            dataGridView1.Sort(this.dataGridView1.Columns[9], ListSortDirection.Ascending);
            dataGridView1.Sort(this.dataGridView1.Columns[1], ListSortDirection.Ascending);
            List<DateTime> data = new List<DateTime>();
            for (int a = 0; a < dataGridView1.Rows.Count - 1; a++)
            {
                for (int b = 0; b < dataGridView1.ColumnCount - 1; b++)
                {
                    //remover lixo da lista como por exemplo espaços
                    if (b == 0)
                    {
                        dataGridView1.Rows[a].Cells[0].Value = formpai.fase1000;
                    }
                    else if (b == 3)
                    {
                        dataGridView1.Rows[a].Cells[3].Value = "2." + lbl_numeroobra.Text + "." + dataGridView1.Rows[a].Cells[0].Value + "." + (a + 1);
                    }
                    else if (b == 4)
                    {
                        dataGridView1.Rows[a].Cells[4].Value = "2." + lbl_numeroobra.Text + "." + dataGridView1.Rows[a].Cells[1].Value + "." + formpai.fase1000 + "H" + (a + 1);
                    }
                    else if (b == 18)
                    {
                        try
                        {
                            dataGridView1.Rows[a].Cells[b].Value = dataGridView1.Rows[a].Cells[b].Value.ToString().Replace(".", "/").Replace("-", "/").Replace("_", "/").Trim();
                            data.Add(Convert.ToDateTime(dataGridView1.Rows[a].Cells[b].Value.ToString()));
                        }
                        catch (Exception)
                        {


                        }

                    }
                    else
                    {
                        dataGridView1.Rows[a].Cells[b].Value = dataGridView1.Rows[a].Cells[b].Value.ToString().Trim();
                    }
                    ///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
                }
            }
            dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
            dataGridView1.Columns[18].DefaultCellStyle.Format = "dd/MM/yyyy";
            data.Sort();
            try
            {
                dateTimePicker1.Text = data[0].ToShortDateString();
            }
            catch (Exception)
            {


            }
            //RESINA();
            RemoverPorcasAnilhasTabela1();
            MudarColunaClasseTabela1();
            DividirValorAnilhasCunhaeMolaTabela1();
            AdicionarAnilhaCunhaouMolaTabela1();
            RemoverEN15048_AnilhasePorcaTabela1();
            RemoverEN14399_AnilhasePorcaTabela1();
            DividirValorColunaEPDMTabela1();
            RemoverBUM_AnilhasePorcaTabela1();
            EmcasodeerrodeReqEspecialTabela1();
            SubstituirPorcasEAnilhasSoltas1();
            ConsolidarLinhasDuplicadas1();
            AjustarColunasDataGridView1();
        }

        public void carregadadosTabela2()
        {
            dataGridView2.Rows.Clear();
            string line = null;
            int i = 1;
            StreamReader file = new StreamReader(@"c:\r\OFELIZ2.CSV", Encoding.Default, true);
            while ((line = file.ReadLine()) != null)
            {
                if (i == 2)
                {
                    var fields = line.Split(';');
                    label1.Text = fields[1];
                }
                if (i == 3)
                {
                    var fields = line.Split(';');
                    label2.Text = fields[1];
                }
                if (i == 4)
                {
                    var fields = line.Split(';');
                    lbl_numeroobra.Text = fields[1];
                }
                if (i == 5)
                {
                    var fields = line.Split(';');
                    label4.Text = fields[1];
                }
                if (i == 6)
                {
                    var fields = line.Split(';');
                    label5.Text = fields[1];
                }

                if (i > 8)
                {
                    var fields = line.Split(';');
                    dataGridView2.Rows.Add(fields);
                }
                i++;
            }
            file.Close();
            File.Delete(@"C:\R\OFELIZ2.CSV");
            dataGridView2.Sort(this.dataGridView2.Columns[9], ListSortDirection.Ascending);
            dataGridView2.Sort(this.dataGridView2.Columns[1], ListSortDirection.Ascending);
            List<DateTime> data = new List<DateTime>();
            LoteeDataemObraTabela2();

            for (int a = 0; a < dataGridView2.Rows.Count - 1; a++)
            {
                for (int b = 0; b < dataGridView2.ColumnCount - 1; b++)
                {
                    //remover lixo da lista como por exemplo espaços
                    if (b == 0)
                    {
                        dataGridView2.Rows[a].Cells[0].Value = formpai.fase1000;
                    }
                    else if (b == 3)
                    {
                        dataGridView2.Rows[a].Cells[3].Value = "2." + lbl_numeroobra.Text + "." + dataGridView2.Rows[a].Cells[0].Value + "." + (a + 1);
                    }
                    else if (b == 4)
                    {
                        dataGridView2.Rows[a].Cells[4].Value = "2." + lbl_numeroobra.Text + "." + dataGridView2.Rows[a].Cells[1].Value + "." + formpai.fase1000 + "H" + (a + 1);
                    }
                    else if (b == 18)
                    {
                        try
                        {
                            dataGridView2.Rows[a].Cells[b].Value = dataGridView2.Rows[a].Cells[b].Value.ToString().Replace(".", "/").Replace("-", "/").Replace("_", "/").Trim();
                            data.Add(Convert.ToDateTime(dataGridView2.Rows[a].Cells[b].Value.ToString()));
                        }
                        catch (Exception)
                        {
                        }

                    }
                    else
                    {
                        dataGridView2.Rows[a].Cells[b].Value = dataGridView2.Rows[a].Cells[b].Value.ToString().Trim();
                    }
                    ///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
                }
            }
            dataGridView2.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
            dataGridView2.Columns[18].DefaultCellStyle.Format = "dd/MM/yyyy";
            data.Sort();
            try
            {
                dateTimePicker1.Text = data[0].ToShortDateString();
            }
            catch (Exception)
            {


            }

            RemoverPorcasAnilhasTabela2();
            MudarColunaClasseTabela2();
            DividirValorAnilhasCunhaeMolaTabela2();
            AdicionarAnilhaCunhaouMolaTabela2();
            RemoverEN15048_AnilhasePorcaTabela2();
            RemoverEN14399_AnilhasePorcaTabela2();
            DividirValorColunaEPDMTabela2();
            RemoverBUM_AnilhasePorcaTabela2();
            EmcasodeerrodeReqEspecialTabela2();
            SubstituirPorcasEAnilhasSoltas2();
            ConsolidarLinhasDuplicadas2();
            ExtrairValores();
            AjustarColunasDataGridView2();
        }

        // Codigo Rui Parafusaria

        //private void RESINA()
        //{
        //    foreach (DataGridViewRow item in dataGridView1.Rows)
        //    {
        //        DataGridViewCell currentcell = item.Cells[47];
        //        if (currentcell.Value!=null)
        //        {
        //            if (currentcell.Value.ToString() != "")
        //            {
        //                int DIAMETRO = int.Parse(item.Cells[9].Value.ToString().Replace("D", ""));

        //                string ArtigoVarao = currentcell.Value.ToString().Split('#')[2] + DIAMETRO;
        //                if (currentcell.Value.ToString().Split('#')[0] == "QUIMICA")
        //                {
        //                    ArtigoVarao = "VRSM" + DIAMETRO;
        //                }

        //                item.Cells[9].Value = ArtigoVarao;

        //                //item.Cells[11].Value = "DIN975";
        //                //item.Cells[12].Value = "2.1";

        //                item.Cells[19].Value = ArtigoVarao;
        //                item.Cells[10].Value = currentcell.Value.ToString().Split('#')[2];
        //                item.Cells[13].Value = (double.Parse(currentcell.Value.ToString().Split('#')[3]) + double.Parse(currentcell.Value.ToString().Split('#')[4])).ToString("0");

        //                int DIAMETROFURO = 0;

        //                if (DIAMETRO < 24)
        //                {
        //                    DIAMETROFURO = DIAMETRO + 2;
        //                }
        //                else
        //                {
        //                    DIAMETROFURO = DIAMETRO + 4;
        //                }




        //                double CALCULO = (Math.PI * Math.Pow(((DIAMETROFURO * 0.01) / 2), 2)) * ((double)2 / (double)3) * (double.Parse(currentcell.Value.ToString().Split('#')[4])) * 0.01 * double.Parse(item.Cells[8].Value.ToString());

        //                List<string> cell = new List<string>();
        //                foreach (DataGridViewCell celula in item.Cells)
        //                {
        //                    cell.Add("" + celula.Value);
        //                }
        //                dataGridView1.Rows.Add(cell.ToArray());
        //                dataGridView1.Rows[dataGridView1.Rows.Count - 2].Cells[8].Value = (CALCULO * 1000).ToString("0");
        //                dataGridView1.Rows[dataGridView1.Rows.Count - 2].Cells[10].Value = currentcell.Value.ToString().Split('#')[1];
        //                currentcell.Value = "";
        //                dataGridView1.Rows[dataGridView1.Rows.Count - 2].Cells[47].Value = "";
        //                dataGridView1.Refresh();

        //            }
        //        }
        //    }
        //}

        /* private void alteraQuantidades()
         {
             foreach (DataGridViewRow item in dataGridView1.Rows)
             {
                 DataGridViewCell currentcell = item.Cells[8];
                 if (currentcell.Value!=null)
                 {
                     if (double.Parse(currentcell.Value.ToString()) <= 150)
                     {
                         currentcell.Value = int.Parse((double.Parse(currentcell.Value.ToString()) + 5).ToString("0"));

                     }
                     else if (double.Parse(currentcell.Value.ToString()) <= 1000)
                     {

                         currentcell.Value = int.Parse((double.Parse(currentcell.Value.ToString()) * ((5.0 / 100.0) + 1)).ToString("0"));

                     }
                     else if (double.Parse(currentcell.Value.ToString()) <= 10000)
                     {
                         currentcell.Value = int.Parse((double.Parse(currentcell.Value.ToString()) * ((2.5 / 100.0) + 1)).ToString("0"));

                     }
                     else
                     {
                         currentcell.Value = int.Parse((double.Parse(currentcell.Value.ToString()) * ((1 / 100.0) + 1)).ToString("0"));

                     }

                     item.Cells[22].Value = "Opção 8";
                     item.Cells[20].Value = "08";
                 }



             }
         }*/

        private void DividirValorAnilhasCunhaeMolaTabela1()
        {
            foreach (DataGridViewRow item in dataGridView1.Rows)
            {
                DataGridViewCell currentCell = item.Cells[11];
                DataGridViewCell valorColuna9 = item.Cells[8];

                if (currentCell.Value != null &&
                    (currentCell.Value.ToString().Contains("DIN 435") ||
                     currentCell.Value.ToString().Contains("DIN 434") ||
                     currentCell.Value.ToString().Contains("DIN 127")))
                {
                    if (valorColuna9.Value != null)
                    {
                        valorColuna9.Value = Convert.ToDouble(valorColuna9.Value) / 2;
                    }
                }
            }
        }

        private string ExtrairDiametro(string valor)
        {
            var match = System.Text.RegularExpressions.Regex.Match(valor, @"\d+");
            if (match.Success)
            {
                return match.Value;
            }
            return string.Empty;
        }

        private void AdicionarAnilhaCunhaouMolaTabela1()
        {
            foreach (DataGridViewRow item in dataGridView1.Rows)
            {
                DataGridViewCell currentCell = item.Cells[11];
                DataGridViewCell classCell = item.Cells[10];
                DataGridViewCell quantidadeCell = item.Cells[8];
                DataGridViewCell diametroCell = item.Cells[9];

                if (currentCell.Value != null &&
                    (currentCell.Value.ToString().Contains("DIN 435") ||
                     currentCell.Value.ToString().Contains("DIN 434") ||
                     currentCell.Value.ToString().Contains("DIN 127")))
                {
                    string classe = classCell.Value != null ? classCell.Value.ToString() : string.Empty;
                    int quantidade = quantidadeCell.Value != null ? Convert.ToInt32(quantidadeCell.Value) : 0;

                    string diametro = ExtrairDiametro(diametroCell.Value.ToString());

                    string diametroatual = "WM" + diametro;

                    bool existeLinha = false;

                    foreach (DataGridViewRow row in dataGridView1.Rows)
                    {
                        if (row.Cells[11].Value != null && row.Cells[11].Value.ToString() == "ISO 7089" &&
                            row.Cells[10].Value != null && row.Cells[10].Value.ToString() == classe &&
                            row.Cells[9].Value != null && row.Cells[9].Value.ToString() == diametroatual)
                        {
                            int quantidadeExistente = row.Cells[8].Value != null ? Convert.ToInt32(row.Cells[8].Value) : 0;
                            row.Cells[8].Value = quantidadeExistente + quantidade;
                            existeLinha = true;
                            break;
                        }
                    }

                    if (!existeLinha)
                    {
                        dataGridView1.Rows.Add(item.Cells[0].Value, item.Cells[1].Value, item.Cells[2].Value, item.Cells[3].Value, item.Cells[4].Value, item.Cells[5].Value, item.Cells[6].Value, item.Cells[7].Value, quantidade, item.Cells[9].Value, classe, "ISO 7089", item.Cells[12].Value, item.Cells[13].Value, item.Cells[14].Value, item.Cells[15].Value, item.Cells[16].Value, item.Cells[17].Value, item.Cells[18].Value, item.Cells[19].Value, item.Cells[20].Value, item.Cells[21].Value, item.Cells[22].Value, item.Cells[23].Value, item.Cells[24].Value);
                    }
                }
            }
        }

        //private void RemoverEN14399AnilhasePorca()
        // {
        //     List<int> quantidadesEN14399 = new List<int>();

        //     foreach (DataGridViewRow item in dataGridView1.Rows)
        //     {
        //         DataGridViewCell currentCell = item.Cells[11];
        //         DataGridViewCell quantidadeCell = item.Cells[8];

        //         if (currentCell.Value != null && currentCell.Value.ToString().Contains("EN-14399"))
        //         {
        //             int quantidade = quantidadeCell.Value != null ? Convert.ToInt32(quantidadeCell.Value) : 0;
        //             quantidadesEN14399.Add(quantidade);
        //         }
        //     }

        //     if (quantidadesEN14399.Count > 0)
        //     {
        //         List<DataGridViewRow> rowsToRemove = new List<DataGridViewRow>();

        //         foreach (DataGridViewRow item in dataGridView1.Rows)
        //         {
        //             DataGridViewCell currentCell = item.Cells[11];
        //             DataGridViewCell quantidadeCell = item.Cells[8];
        //             DataGridViewCell diametroCell = item.Cells[9];

        //             string diametro = ExtrairDiametro(diametroCell.Value?.ToString() ?? string.Empty);
        //             string diametroatual = "NM" + diametro;

        //             if (currentCell.Value != null &&
        //                 currentCell.Value.ToString().Contains("ISO 4032") &&
        //                 item.Cells[10].Value != null && item.Cells[10].Value.ToString().Contains("10 ZINCADO") &&
        //                 item.Cells[9].Value != null && item.Cells[9].Value.ToString() == diametroatual)
        //             {
        //                 int quantidadeExistente = quantidadeCell.Value != null ? Convert.ToInt32(quantidadeCell.Value) : 0;

        //                 foreach (int quantidadeEN14399 in quantidadesEN14399)
        //                 {
        //                     quantidadeExistente -= quantidadeEN14399;
        //                 }

        //                 quantidadeCell.Value = Math.Max(0, quantidadeExistente);

        //                 if (quantidadeCell.Value.Equals(0))
        //                 {
        //                     rowsToRemove.Add(item);
        //                 }
        //             }
        //         }

        //         foreach (var row in rowsToRemove)
        //         {
        //             dataGridView1.Rows.Remove(row);
        //         }

        //         rowsToRemove.Clear(); 

        //         foreach (DataGridViewRow item in dataGridView1.Rows)
        //         {
        //             DataGridViewCell currentCell = item.Cells[11];
        //             DataGridViewCell quantidadeCell = item.Cells[8];
        //             DataGridViewCell diametroCell = item.Cells[9];

        //             string diametro = ExtrairDiametro(diametroCell.Value?.ToString() ?? string.Empty);
        //             string diametroatual = "WM" + diametro;

        //             if (currentCell.Value != null &&
        //                 currentCell.Value.ToString().Contains("ISO 7089") &&
        //                 item.Cells[10].Value != null && item.Cells[10].Value.ToString().Contains("300 HV ZINCADO") &&
        //                 item.Cells[9].Value != null && item.Cells[9].Value.ToString() == diametroatual)
        //             {
        //                 int quantidadeExistente = quantidadeCell.Value != null ? Convert.ToInt32(quantidadeCell.Value) : 0;

        //                 foreach (int quantidadeEN14399 in quantidadesEN14399)
        //                 {
        //                     quantidadeExistente -= (quantidadeEN14399 * 2);
        //                 }

        //                 quantidadeCell.Value = Math.Max(0, quantidadeExistente);

        //                 if (quantidadeCell.Value.Equals(0))
        //                 {
        //                     rowsToRemove.Add(item);
        //                 }
        //             }
        //         }

        //         foreach (var row in rowsToRemove)
        //         {
        //             dataGridView1.Rows.Remove(row);
        //         }
        //     }
        // }

        private void RemoverEN14399_AnilhasePorcaTabela1()
        {
            List<int> quantidadesEN14399 = new List<int>();
            List<string> diametrosEN14399 = new List<string>();

            foreach (DataGridViewRow item in dataGridView1.Rows)
            {
                DataGridViewCell currentCell = item.Cells[11];
                DataGridViewCell quantidadeCell = item.Cells[8];
                DataGridViewCell diametroCell = item.Cells[9];

                if (currentCell.Value != null && currentCell.Value.ToString().Contains("EN-14399"))
                {
                    int quantidade = quantidadeCell.Value != null ? Convert.ToInt32(quantidadeCell.Value) : 0;
                    quantidadesEN14399.Add(quantidade);
                    diametrosEN14399.Add(ExtrairDiametro(diametroCell.Value?.ToString() ?? string.Empty));
                }
            }

            if (quantidadesEN14399.Count > 0)
            {
                List<DataGridViewRow> rowsToRemove = new List<DataGridViewRow>();

                foreach (DataGridViewRow item in dataGridView1.Rows)
                {
                    DataGridViewCell currentCell = item.Cells[11];
                    DataGridViewCell quantidadeCell = item.Cells[8];
                    DataGridViewCell diametroCell = item.Cells[9];

                    string diametro = ExtrairDiametro(diametroCell.Value?.ToString() ?? string.Empty);
                    string diametroatual = "NM" + diametro;

                    if (currentCell.Value != null &&
                        currentCell.Value.ToString().Contains("ISO 4032") &&
                        item.Cells[10].Value != null && item.Cells[10].Value.ToString().Contains("10") &&
                        item.Cells[9].Value != null && item.Cells[9].Value.ToString() == diametroatual)
                    {
                        int quantidadeExistente = quantidadeCell.Value != null ? Convert.ToInt32(quantidadeCell.Value) : 0;

                        for (int i = 0; i < quantidadesEN14399.Count; i++)
                        {
                            if (diametrosEN14399[i] == diametro)
                            {
                                quantidadeExistente -= quantidadesEN14399[i];
                            }
                        }

                        quantidadeCell.Value = Math.Max(0, quantidadeExistente);

                        if (quantidadeCell.Value.Equals(0))
                        {
                            rowsToRemove.Add(item);
                        }
                    }
                }

                foreach (var row in rowsToRemove)
                {
                    dataGridView1.Rows.Remove(row);
                }

                rowsToRemove.Clear();

                foreach (DataGridViewRow item in dataGridView1.Rows)
                {
                    DataGridViewCell currentCell = item.Cells[11];
                    DataGridViewCell quantidadeCell = item.Cells[8];
                    DataGridViewCell diametroCell = item.Cells[9];

                    string diametro = ExtrairDiametro(diametroCell.Value?.ToString() ?? string.Empty);
                    string diametroatual = "WM" + diametro;

                    if (currentCell.Value != null &&
                        currentCell.Value.ToString().Contains("ISO 7089") &&
                        item.Cells[10].Value != null && item.Cells[10].Value.ToString().Contains("300 HV") &&
                        item.Cells[9].Value != null && item.Cells[9].Value.ToString() == diametroatual)
                    {
                        int quantidadeExistente = quantidadeCell.Value != null ? Convert.ToInt32(quantidadeCell.Value) : 0;

                        for (int i = 0; i < quantidadesEN14399.Count; i++)
                        {
                            if (diametrosEN14399[i] == diametro)
                            {
                                quantidadeExistente -= (quantidadesEN14399[i] * 2);
                            }
                        }

                        quantidadeCell.Value = Math.Max(0, quantidadeExistente);

                        if (quantidadeCell.Value.Equals(0))
                        {
                            rowsToRemove.Add(item);
                        }
                    }
                }

                foreach (var row in rowsToRemove)
                {
                    dataGridView1.Rows.Remove(row);
                }
            }
        }

        private void RemoverEN14399_AnilhasePorcaTabela2()
        {
            List<int> quantidadesEN14399 = new List<int>();
            List<string> diametrosEN14399 = new List<string>();

            foreach (DataGridViewRow item in dataGridView2.Rows)
            {
                DataGridViewCell currentCell = item.Cells[11];
                DataGridViewCell quantidadeCell = item.Cells[8];
                DataGridViewCell diametroCell = item.Cells[9];

                if (currentCell.Value != null && currentCell.Value.ToString().Contains("EN-14399"))
                {
                    int quantidade = quantidadeCell.Value != null ? Convert.ToInt32(quantidadeCell.Value) : 0;
                    quantidadesEN14399.Add(quantidade);
                    diametrosEN14399.Add(ExtrairDiametro(diametroCell.Value?.ToString() ?? string.Empty));
                }
            }

            if (quantidadesEN14399.Count > 0)
            {
                List<DataGridViewRow> rowsToRemove = new List<DataGridViewRow>();

                foreach (DataGridViewRow item in dataGridView2.Rows)
                {
                    DataGridViewCell currentCell = item.Cells[11];
                    DataGridViewCell quantidadeCell = item.Cells[8];
                    DataGridViewCell diametroCell = item.Cells[9];

                    string diametro = ExtrairDiametro(diametroCell.Value?.ToString() ?? string.Empty);
                    string diametroatual = "NM" + diametro;

                    if (currentCell.Value != null &&
                        currentCell.Value.ToString().Contains("ISO 4032") &&
                        item.Cells[10].Value != null && item.Cells[10].Value.ToString().Contains("10") &&
                        item.Cells[9].Value != null && item.Cells[9].Value.ToString() == diametroatual)
                    {
                        int quantidadeExistente = quantidadeCell.Value != null ? Convert.ToInt32(quantidadeCell.Value) : 0;

                        for (int i = 0; i < quantidadesEN14399.Count; i++)
                        {
                            if (diametrosEN14399[i] == diametro)
                            {
                                quantidadeExistente -= quantidadesEN14399[i];
                            }
                        }

                        quantidadeCell.Value = Math.Max(0, quantidadeExistente);

                        if (quantidadeCell.Value.Equals(0))
                        {
                            rowsToRemove.Add(item);
                        }
                    }
                }

                foreach (var row in rowsToRemove)
                {
                    dataGridView2.Rows.Remove(row);
                }

                rowsToRemove.Clear();

                foreach (DataGridViewRow item in dataGridView2.Rows)
                {
                    DataGridViewCell currentCell = item.Cells[11];
                    DataGridViewCell quantidadeCell = item.Cells[8];
                    DataGridViewCell diametroCell = item.Cells[9];

                    string diametro = ExtrairDiametro(diametroCell.Value?.ToString() ?? string.Empty);
                    string diametroatual = "WM" + diametro;

                    if (currentCell.Value != null &&
                        currentCell.Value.ToString().Contains("ISO 7089") &&
                        item.Cells[10].Value != null && item.Cells[10].Value.ToString().Contains("300 HV") &&
                        item.Cells[9].Value != null && item.Cells[9].Value.ToString() == diametroatual)
                    {
                        int quantidadeExistente = quantidadeCell.Value != null ? Convert.ToInt32(quantidadeCell.Value) : 0;

                        for (int i = 0; i < quantidadesEN14399.Count; i++)
                        {
                            if (diametrosEN14399[i] == diametro)
                            {
                                quantidadeExistente -= (quantidadesEN14399[i] * 2);
                            }
                        }

                        quantidadeCell.Value = Math.Max(0, quantidadeExistente);

                        if (quantidadeCell.Value.Equals(0))
                        {
                            rowsToRemove.Add(item);
                        }
                    }
                }


                foreach (var row in rowsToRemove)
                {
                    dataGridView2.Rows.Remove(row);
                }
            }
        }

        private void RemoverPorcasAnilhasTabela1()
        {
            List<DataGridViewRow> rowsToRemove = new List<DataGridViewRow>();

            foreach (DataGridViewRow item in dataGridView1.Rows)
            {
                DataGridViewCell currentcell = item.Cells[11];

                if (currentcell.Value != null && currentcell.Value.ToString().Contains("Remover"))
                {
                    rowsToRemove.Add(item);
                }
            }

            foreach (DataGridViewRow row in rowsToRemove)
            {
                dataGridView1.Rows.Remove(row);
            }
        }

        private void RemoverPorcasAnilhasTabela2()
        {
            List<DataGridViewRow> rowsToRemove = new List<DataGridViewRow>();

            foreach (DataGridViewRow item in dataGridView2.Rows)
            {
                DataGridViewCell currentcell = item.Cells[11];

                if (currentcell.Value != null && currentcell.Value.ToString().Contains("Remover"))
                {
                    rowsToRemove.Add(item);
                }
            }

            foreach (DataGridViewRow row in rowsToRemove)
            {
                dataGridView2.Rows.Remove(row);
            }
        }

        private void RemoverEN15048_AnilhasePorcaTabela1()
        {
            Dictionary<string, string> mapeamentoClasseNW = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
                     {
                        { "8.8 ZINCADO", "8 ZINCADO" },
                        { "8.8-SB ZINCADO", "8 ZINCADO" },
                        { "10.9-HV ZINCADO", "10 ZINCADO" },
                        { "8,8 ZINCADO 2,1", "8 ZINCADO" },
                        { "8,8 ZINCADO 2,2", "8 ZINCADO" },
                        { "8,8 ZINCADO 3,1", "8 ZINCADO" },
                        { "8,8 ZINCADO 3,2", "8 ZINCADO" },
                        { "10,9 ZINCADO 2,1", "10 ZINCADO" },
                        { "10,9 ZINCADO 2,2", "10 ZINCADO" },
                        { "10,9 ZINCADO 3,1", "10 ZINCADO" },
                        { "10,9 ZINCADO 3,2", "10 ZINCADO" },

                        { "8.8-SB PRETO", "8 PRETO" },
                        { "10.9-HV PRETO", "10 PRETO" },
                        { "8,8 PRETO 2,1", "8 PRETO" },
                        { "8,8 PRETO 2,2", "8 PRETO" },
                        { "8,8 PRETO 3,1", "8 PRETO" },
                        { "8,8 PRETO 3,2", "8 PRETO" },
                        { "10,9 PRETO 2,1", "10 PRETO" },
                        { "10,9 PRETO 2,2", "10 PRETO" },
                        { "10,9 PRETO 3,1", "10 PRETO" },
                        { "10,9 PRETO 3,2", "10 PRETO" },

                        { "8.8-SB GALVANIZADO", "8 GALVANIZADO" },
                        { "10.9-HV GALVANIZADO", "10 GALVANIZADO" },
                        { "8,8 GALVANIZADO 2,1", "8 GALVANIZADO" },
                        { "8,8 GALVANIZADO 2,2", "8 GALVANIZADO" },
                        { "8,8 GALVANIZADO 3,1", "8 GALVANIZADO" },
                        { "8,8 GALVANIZADO 3,2", "8 GALVANIZADO" },
                        { "10,9 GALVANIZADO 2,1", "10 GALVANIZADO" },
                        { "10,9 GALVANIZADO 2,2", "10 GALVANIZADO" },
                        { "10,9 GALVANIZADO 3,1", "10 GALVANIZADO" },
                        { "10,9 GALVANIZADO 3,2", "10 GALVANIZADO" }
                     };


            List<int> quantidadesEN15048 = new List<int>();
            List<string> diametrosEN15048 = new List<string>();
            List<string> classesPorca = new List<string>();

            foreach (DataGridViewRow item in dataGridView1.Rows)
            {
                DataGridViewCell currentCell = item.Cells[11];
                DataGridViewCell quantidadeCell = item.Cells[8];
                DataGridViewCell diametroCell = item.Cells[9];
                DataGridViewCell classeCell = item.Cells[10];

                if (currentCell.Value != null && currentCell.Value.ToString().Contains("EN15048"))
                {
                    int quantidade = quantidadeCell.Value != null ? Convert.ToInt32(quantidadeCell.Value) : 0;
                    string classeOriginal = classeCell.Value?.ToString();

                    string classePorca = "";

                    if (!string.IsNullOrEmpty(classeOriginal))
                    {
                        mapeamentoClasseNW.TryGetValue(classeOriginal, out classePorca);
                    }

                    string diametro = ExtrairDiametro(diametroCell.Value?.ToString() ?? string.Empty);
                    string diametroNM = "NM" + diametro;

                    quantidadesEN15048.Add(quantidade);
                    diametrosEN15048.Add(diametroNM);
                    classesPorca.Add(classePorca);
                }
            }

            if (quantidadesEN15048.Count > 0)
            {
                List<DataGridViewRow> rowsToRemove = new List<DataGridViewRow>();

                foreach (DataGridViewRow item in dataGridView1.Rows)
                {
                    DataGridViewCell currentCell = item.Cells[11];
                    DataGridViewCell quantidadeCell = item.Cells[8];
                    DataGridViewCell diametroCell = item.Cells[9];
                    DataGridViewCell classeCell = item.Cells[10];

                    string diametro = ExtrairDiametro(diametroCell.Value?.ToString() ?? string.Empty);
                    string diametroNM = "NM" + diametro;

                    if (currentCell.Value != null &&
                        currentCell.Value.ToString().Contains("ISO 4032") &&
                        classeCell.Value != null &&
                        item.Cells[9].Value != null &&
                        item.Cells[9].Value.ToString() == diametroNM)
                    {
                        int quantidadeExistente = quantidadeCell.Value != null ? Convert.ToInt32(quantidadeCell.Value) : 0;
                        string classe = classeCell.Value.ToString();

                        for (int i = 0; i < quantidadesEN15048.Count; i++)
                        {
                            if (diametrosEN15048[i] == diametroNM && classesPorca[i] == classe)
                            {
                                quantidadeExistente -= quantidadesEN15048[i];
                            }
                        }

                        quantidadeCell.Value = Math.Max(0, quantidadeExistente);
                        if (quantidadeCell.Value.Equals(0))
                        {
                            rowsToRemove.Add(item);
                        }
                    }
                }

                foreach (var row in rowsToRemove)
                    dataGridView1.Rows.Remove(row);
            }
        }

        private void RemoverEN15048_AnilhasePorcaTabela2()
        {
            Dictionary<string, string> mapeamentoClasseNW = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
                     {
                        { "8.8 ZINCADO", "8 ZINCADO" },
                        { "8.8-SB ZINCADO", "8 ZINCADO" },
                        { "10.9-HV ZINCADO", "10 ZINCADO" },
                        { "8,8 ZINCADO 2,1", "8 ZINCADO" },
                        { "8,8 ZINCADO 2,2", "8 ZINCADO" },
                        { "8,8 ZINCADO 3,1", "8 ZINCADO" },
                        { "8,8 ZINCADO 3,2", "8 ZINCADO" },
                        { "10,9 ZINCADO 2,1", "10 ZINCADO" },
                        { "10,9 ZINCADO 2,2", "10 ZINCADO" },
                        { "10,9 ZINCADO 3,1", "10 ZINCADO" },
                        { "10,9 ZINCADO 3,2", "10 ZINCADO" },

                        { "8.8-SB PRETO", "8 PRETO" },
                        { "10.9-HV PRETO", "10 PRETO" },
                        { "8,8 PRETO 2,1", "8 PRETO" },
                        { "8,8 PRETO 2,2", "8 PRETO" },
                        { "8,8 PRETO 3,1", "8 PRETO" },
                        { "8,8 PRETO 3,2", "8 PRETO" },
                        { "10,9 PRETO 2,1", "10 PRETO" },
                        { "10,9 PRETO 2,2", "10 PRETO" },
                        { "10,9 PRETO 3,1", "10 PRETO" },
                        { "10,9 PRETO 3,2", "10 PRETO" },

                        { "8.8-SB GALVANIZADO", "8 GALVANIZADO" },
                        { "10.9-HV GALVANIZADO", "10 GALVANIZADO" },
                        { "8,8 GALVANIZADO 2,1", "8 GALVANIZADO" },
                        { "8,8 GALVANIZADO 2,2", "8 GALVANIZADO" },
                        { "8,8 GALVANIZADO 3,1", "8 GALVANIZADO" },
                        { "8,8 GALVANIZADO 3,2", "8 GALVANIZADO" },
                        { "10,9 GALVANIZADO 2,1", "10 GALVANIZADO" },
                        { "10,9 GALVANIZADO 2,2", "10 GALVANIZADO" },
                        { "10,9 GALVANIZADO 3,1", "10 GALVANIZADO" },
                        { "10,9 GALVANIZADO 3,2", "10 GALVANIZADO" }
                     };


            //    List<int> quantidadesEN15048 = new List<int>();
            //    List<string> diametrosEN15048 = new List<string>();
            //    List<string> classesPorca = new List<string>();

            //    foreach (DataGridViewRow item in dataGridView2.Rows)
            //    {
            //        DataGridViewCell currentCell = item.Cells[11];
            //        DataGridViewCell quantidadeCell = item.Cells[8];
            //        DataGridViewCell diametroCell = item.Cells[9];
            //        DataGridViewCell classeCell = item.Cells[10];

            //        if (currentCell.Value != null && currentCell.Value.ToString().Contains("EN15048"))
            //        {
            //            int quantidade = quantidadeCell.Value != null ? Convert.ToInt32(quantidadeCell.Value) : 0;
            //            string classeOriginal = classeCell.Value?.ToString();

            //            string classePorca = "";

            //            if (!string.IsNullOrEmpty(classeOriginal))
            //            {
            //                mapeamentoClasseNW.TryGetValue(classeOriginal, out classePorca);
            //            }

            //            string diametro = ExtrairDiametro(diametroCell.Value?.ToString() ?? string.Empty);
            //            string diametroNM = "NM" + diametro;

            //            quantidadesEN15048.Add(quantidade);
            //            diametrosEN15048.Add(diametroNM);
            //            classesPorca.Add(classePorca);
            //        }
            //    }

            //    if (quantidadesEN15048.Count > 0)
            //    {
            //        List<DataGridViewRow> rowsToRemove = new List<DataGridViewRow>();

            //        foreach (DataGridViewRow item in dataGridView2.Rows)
            //        {
            //            DataGridViewCell currentCell = item.Cells[11];
            //            DataGridViewCell quantidadeCell = item.Cells[8];
            //            DataGridViewCell diametroCell = item.Cells[9];
            //            DataGridViewCell classeCell = item.Cells[10];

            //            string diametro = ExtrairDiametro(diametroCell.Value?.ToString() ?? string.Empty);
            //            string diametroNM = "NM" + diametro;

            //            if (currentCell.Value != null &&
            //                currentCell.Value.ToString().Contains("ISO 4032") &&
            //                classeCell.Value != null &&
            //                item.Cells[9].Value != null &&
            //                item.Cells[9].Value.ToString() == diametroNM)
            //            {
            //                int quantidadeExistente = quantidadeCell.Value != null ? Convert.ToInt32(quantidadeCell.Value) : 0;
            //                string classe = classeCell.Value.ToString();

            //                for (int i = 0; i < quantidadesEN15048.Count; i++)
            //                {
            //                    if (diametrosEN15048[i] == diametroNM && classesPorca[i] == classe)
            //                    {
            //                        quantidadeExistente -= quantidadesEN15048[i];
            //                    }
            //                }

            //                quantidadeCell.Value = Math.Max(0, quantidadeExistente);
            //                if (quantidadeCell.Value.Equals(0))
            //                {
            //                    rowsToRemove.Add(item);
            //                }
            //            }
            //        }

            //        foreach (var row in rowsToRemove)
            //            dataGridView2.Rows.Remove(row);
            //    }
            //}

            Dictionary<(string diametro, string classe), int> totalPorcasEN15048 = new Dictionary<(string diametro, string classe), int>();

            // Primeira passagem: identificar os parafusos EN15048 e acumular
            foreach (DataGridViewRow item in dataGridView2.Rows)
            {
                var cellTipo = item.Cells[11];
                var cellQtd = item.Cells[8];
                var cellDiametro = item.Cells[9];
                var cellClasse = item.Cells[10];

                if (cellTipo?.Value?.ToString().Contains("EN15048") == true)
                {
                    int quantidade = cellQtd?.Value != null ? Convert.ToInt32(cellQtd.Value) : 0;
                    string classeOriginal = cellClasse?.Value?.ToString()?.Trim() ?? "";

                    if (mapeamentoClasseNW.TryGetValue(classeOriginal, out string classePorca))
                    {
                        string diametro = "NM" + ExtrairDiametro(cellDiametro?.Value?.ToString() ?? "");

                        var chave = (diametro, classePorca);
                        if (totalPorcasEN15048.ContainsKey(chave))
                            totalPorcasEN15048[chave] += quantidade;
                        else
                            totalPorcasEN15048[chave] = quantidade;
                    }
                }
            }

            // Lista de linhas para remoção
            List<DataGridViewRow> rowsToRemove = new List<DataGridViewRow>();

            // Segunda passagem: localizar porcas ISO 4032 e subtrair
            foreach (DataGridViewRow item in dataGridView2.Rows)
            {
                var cellTipo = item.Cells[11];
                var cellQtd = item.Cells[8];
                var cellDiametro = item.Cells[9];
                var cellClasse = item.Cells[10];

                if (cellTipo?.Value?.ToString().Contains("ISO 4032") == true)
                {
                    string diametro = "NM" + ExtrairDiametro(cellDiametro?.Value?.ToString() ?? "");
                    string classe = cellClasse?.Value?.ToString()?.Trim() ?? "";

                    var chave = (diametro, classe);
                    if (totalPorcasEN15048.TryGetValue(chave, out int qtdParaSubtrair))
                    {
                        int qtdExistente = cellQtd?.Value != null ? Convert.ToInt32(cellQtd.Value) : 0;
                        cellQtd.Value = Math.Max(0, qtdExistente - qtdParaSubtrair);

                        if ((int)cellQtd.Value == 0)
                            rowsToRemove.Add(item);
                    }
                }
            }

            // Remover linhas com quantidade 0
            foreach (var row in rowsToRemove)
            {
                dataGridView2.Rows.Remove(row);
            }
        }

        private void DividirValorAnilhasCunhaeMolaTabela2()
        {
            foreach (DataGridViewRow item in dataGridView2.Rows)
            {
                DataGridViewCell currentCell = item.Cells[11];
                DataGridViewCell valorColuna9 = item.Cells[8];

                if (currentCell.Value != null &&
                    (currentCell.Value.ToString().Contains("DIN 435") ||
                     currentCell.Value.ToString().Contains("DIN 434") ||
                     currentCell.Value.ToString().Contains("DIN 127")))
                {
                    if (valorColuna9.Value != null)
                    {
                        valorColuna9.Value = Convert.ToDouble(valorColuna9.Value) / 2;
                    }
                }
            }
        }

        private void AdicionarAnilhaCunhaouMolaTabela2()
        {
            foreach (DataGridViewRow item in dataGridView2.Rows)
            {
                DataGridViewCell currentCell = item.Cells[11];
                DataGridViewCell classCell = item.Cells[10];
                DataGridViewCell quantidadeCell = item.Cells[8];
                DataGridViewCell diametroCell = item.Cells[9];

                if (currentCell.Value != null &&
                    (currentCell.Value.ToString().Contains("DIN 435") ||
                     currentCell.Value.ToString().Contains("DIN 434") ||
                     currentCell.Value.ToString().Contains("DIN 127")))
                {
                    string classe = classCell.Value != null ? classCell.Value.ToString() : string.Empty;
                    int quantidade = quantidadeCell.Value != null ? Convert.ToInt32(quantidadeCell.Value) : 0;

                    string diametro = ExtrairDiametro(diametroCell.Value.ToString());

                    string diametroatual = "WM" + diametro;

                    bool existeLinha = false;

                    foreach (DataGridViewRow row in dataGridView2.Rows)
                    {
                        if (row.Cells[11].Value != null && row.Cells[11].Value.ToString() == "ISO 7089" &&
                            row.Cells[10].Value != null && row.Cells[10].Value.ToString() == classe &&
                            row.Cells[9].Value != null && row.Cells[9].Value.ToString() == diametroatual)
                        {
                            int quantidadeExistente = row.Cells[8].Value != null ? Convert.ToInt32(row.Cells[8].Value) : 0;
                            row.Cells[8].Value = quantidadeExistente + quantidade;
                            existeLinha = true;
                            break;
                        }
                    }

                    if (!existeLinha)
                    {
                        dataGridView2.Rows.Add(item.Cells[0].Value, item.Cells[1].Value, item.Cells[2].Value, item.Cells[3].Value, item.Cells[4].Value, item.Cells[5].Value, item.Cells[6].Value, item.Cells[7].Value, quantidade, item.Cells[9].Value, classe, "ISO 7089", item.Cells[12].Value, item.Cells[13].Value, item.Cells[14].Value, item.Cells[15].Value, item.Cells[16].Value, item.Cells[17].Value, item.Cells[18].Value, item.Cells[19].Value, item.Cells[20].Value, item.Cells[21].Value, item.Cells[22].Value, item.Cells[23].Value, item.Cells[24].Value);
                    }
                }
            }
        }

        private void RemoverBUM_AnilhasePorcaTabela1()
        {
            foreach (DataGridViewRow item in dataGridView1.Rows)
            {
                DataGridViewCell coluna9 = item.Cells[9];
                DataGridViewCell coluna10 = item.Cells[10];
                DataGridViewCell coluna13 = item.Cells[13];

                if (coluna10.Value != null)
                {
                    string valorColuna10 = coluna10.Value.ToString().Trim();
                    if (valorColuna10.StartsWith("BUM"))
                    {
                        coluna10.Value = valorColuna10.Substring(3).Trim();
                    }
                }

                if (coluna9.Value != null && coluna9.Value.ToString().Trim().StartsWith("BUM"))
                {
                    if (coluna13.Value == null || string.IsNullOrWhiteSpace(coluna13.Value.ToString()))
                    {
                        coluna13.Value = "100";
                    }
                    else
                    {
                        if (int.TryParse(coluna13.Value.ToString(), out int valor13))
                        {
                            if (valor13 < 99)
                            {
                                coluna13.Value = "100";
                            }
                        }
                        else
                        {
                            coluna13.Value = "100";
                        }
                    }
                }
            }
        }

        private void RemoverBUM_AnilhasePorcaTabela2()
        {
            foreach (DataGridViewRow item in dataGridView2.Rows)
            {
                DataGridViewCell coluna9 = item.Cells[9];
                DataGridViewCell coluna10 = item.Cells[10];
                DataGridViewCell coluna13 = item.Cells[13];

                if (coluna10.Value != null)
                {
                    string valorColuna10 = coluna10.Value.ToString().Trim();
                    if (valorColuna10.StartsWith("BUM"))
                    {
                        coluna10.Value = valorColuna10.Substring(3).Trim();
                    }
                }

                if (coluna9.Value != null && coluna9.Value.ToString().Trim().StartsWith("BUM"))
                {
                    if (coluna13.Value == null || string.IsNullOrWhiteSpace(coluna13.Value.ToString()))
                    {
                        coluna13.Value = "100";
                    }
                    else
                    {
                        if (int.TryParse(coluna13.Value.ToString(), out int valor13))
                        {
                            if (valor13 < 99)
                            {
                                coluna13.Value = "100";
                            }
                        }
                        else
                        {
                            coluna13.Value = "100";
                        }
                    }
                }
            }
        }


        //private void button2_Click(object sender, EventArgs e)
        //{
        //    if (!Directory.Exists(Frm_Inico.CaminhoModelo + @"\listas"))
        //    {
        //        Directory.CreateDirectory(Frm_Inico.CaminhoModelo + @"\listas");
        //    }
        //    string Save = null;
        //    if (Directory.Exists(Frm_Inico.PastaReservatorioFicheiros + formpai.fase1000))
        //    {
        //        Save = Frm_Inico.PastaReservatorioFicheiros + formpai.fase1000 + "\\" + lbl_numeroobra.Text + "F" + formpai.fase1000 + ".csv";
        //    }
        //    else
        //    {
        //        Directory.CreateDirectory(Frm_Inico.PastaReservatorioFicheiros + formpai.fase1000);
        //        Save = Frm_Inico.PastaReservatorioFicheiros + formpai.fase1000 + "\\" + lbl_numeroobra.Text + "F" + formpai.fase1000 + ".csv";
        //    }

        //    if (!Directory.Exists(Frm_Inico.PastaPartilhada+"\\"+Frm_Inico.ano+"\\ARM\\"+ lbl_numeroobra.Text+"\\"+ formpai.fase1000))
        //    {
        //        Directory.CreateDirectory(Frm_Inico.PastaPartilhada + "\\" + Frm_Inico.ano + "\\ARM\\" + lbl_numeroobra.Text + "\\" + formpai.fase1000);
        //    }
        //    if (!Directory.Exists(Frm_Inico.PastaPartilhada + "\\" + Frm_Inico.ano + "\\ARM\\" + lbl_numeroobra.Text + "\\" + formpai.fase1000))
        //    {
        //        Directory.CreateDirectory(Frm_Inico.PastaPartilhada + "\\" + Frm_Inico.ano + "\\ARM\\" + lbl_numeroobra.Text + "\\" + formpai.fase1000);
        //    }
        //    if (!Directory.Exists(Frm_Inico.PastaReservatorioFicheiros + formpai.fase1000+"\\20001"))
        //    {
        //        Directory.CreateDirectory(Frm_Inico.PastaReservatorioFicheiros + formpai.fase1000 + "\\20001");
        //    }
        //    if (!Directory.Exists(Frm_Inico.PastaReservatorioFicheiros + formpai.fase1000 + "\\20009"))
        //    {
        //        Directory.CreateDirectory(Frm_Inico.PastaReservatorioFicheiros + formpai.fase1000 + "\\20009");
        //    }
        //    SaveToCSV(dataGridView1, Save);
        //    Tekla.Structures.Model.Model m = new Tekla.Structures.Model.Model();
        //    string numeroobra = m.GetProjectInfo().ProjectNumber;
        //    MessageBox.Show("Lista exportada " + Environment.NewLine + " Fase:" + formpai.fase1000, "Exportação", MessageBoxButtons.OK, MessageBoxIcon.Information);

        //}

        private void SaveToCSV(DataGridView DGV, string filename)
        {
            int columnCount = DGV.ColumnCount;
            string columnNames = "";
            string[] output = new string[DGV.RowCount + 7];

            for (int i = 0; i < columnCount; i++)
            {
                columnNames += DGV.Columns[i].HeaderText.ToString() + ";";
            }
            output[0] = "O FELIZ FICHA DE PEÇAS";
            output[1] = "Designação:;" + label1.Text;
            output[2] = "Cliente:;" + label2.Text;
            output[3] = "Nº Obra:;" + lbl_numeroobra.Text;
            output[4] = "Data:;" + label4.Text;
            output[5] = "Classe de Execução:;" + label5.Text;
            output[6] = "Observações;;;;;;;;;;;;;;" + label3.Text + ";" + label6.Text;
            output[7] += columnNames;
            int a = 1;
            for (int i = 8; (i - 8) < DGV.RowCount - 1; i++)
            {
                for (int j = 0; j < columnCount; j++)
                {
                    try
                    {
                        if (DGV.Rows[i - 8].DefaultCellStyle.BackColor != Color.Red)
                        {
                            output[i] += DGV.Rows[i - 8].Cells[j].Value.ToString() + ";";
                        }
                        else
                        {
                            if (a == 1)
                            {
                                MessageBox.Show("As linhas a vermelho não foram exportadas", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                a++;
                            }
                        }
                    }
                    catch (Exception)
                    {

                    }
                }
            }
            System.IO.File.WriteAllLines(filename, output, System.Text.Encoding.Default);
        }

        private void SaveToCSV2(DataGridView DGV, string filename)
        {
            int columnCount = DGV.ColumnCount;
            string columnNames = "";
            string[] output = new string[DGV.RowCount + 7];

            for (int i = 0; i < columnCount; i++)
            {
                columnNames += DGV.Columns[i].HeaderText.ToString() + ";";
            }
            output[0] = "O FELIZ FICHA DE PEÇAS";
            output[1] = "Designação:;" + label1.Text;
            output[2] = "Cliente:;" + label2.Text;
            output[3] = "Nº Obra:;" + lbl_numeroobra.Text;
            output[4] = "Data:;" + label4.Text;
            output[5] = "Classe de Execução:;" + label5.Text;
            output[6] = "Observações;;;;;;;;;;;;;;" + label3.Text + ";" + label6.Text;
            output[7] += columnNames;
            int a = 1;
            for (int i = 8; (i - 8) < DGV.RowCount - 1; i++)
            {
                for (int j = 0; j < columnCount; j++)
                {
                    try
                    {
                        if (DGV.Rows[i - 8].DefaultCellStyle.BackColor != Color.Red)
                        {
                            output[i] += DGV.Rows[i - 8].Cells[j].Value.ToString() + ";";
                        }
                        else
                        {
                            if (a == 1)
                            {
                                MessageBox.Show("As linhas a vermelho não foram exportadas", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                a++;
                            }
                        }
                    }
                    catch (Exception)
                    {

                    }
                }
            }
            System.IO.File.WriteAllLines(filename, output, System.Text.Encoding.Default);
        }

        private void LoteeDataemObraTabela2()
        {
            string lote = null;
            string dataObra = null;

            int? menorLote = null;
            DateTime? menorData = null;

            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                if (!row.IsNewRow)
                {
                    if (row.Cells[2].Value != null)
                    {
                        int loteAtual;
                        if (int.TryParse(row.Cells[1].Value.ToString(), out loteAtual))
                        {
                            if (menorLote == null || loteAtual < menorLote)
                            {
                                menorLote = loteAtual;
                            }
                        }
                    }

                    if (row.Cells[18].Value != null)
                    {
                        DateTime dataAtual;

                        if (DateTime.TryParse(row.Cells[18].Value.ToString(), out dataAtual))
                        {
                            if (menorData == null || dataAtual < menorData)
                            {
                                menorData = dataAtual;
                            }
                        }
                    }
                }
            }

            if (menorData != null)
            {
                string menorDataString = menorData.Value.ToString("yyyy-MM-dd");

            }
            else
            {
                string menorDataString = null;
            }

            if (menorLote.HasValue)
            {
                lote = menorLote.Value.ToString();
            }

            if (menorData.HasValue)
            {
                dataObra = menorData.Value.ToString("dd/MM/yyyy");
            }

            if (dataGridView2 != null)
            {
                foreach (DataGridViewRow row in dataGridView2.Rows)
                {
                    if (!row.IsNewRow)
                    {
                        row.Cells[1].Value = lote;

                        row.Cells[18].Value = dataObra;
                    }
                }
            }
        }

        private void DividirValorColunaEPDMTabela1()
        {
            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                if (!row.IsNewRow)
                {
                    if (row.Cells[11].Value != null && row.Cells[11].Value.ToString().Equals("EPDM", StringComparison.OrdinalIgnoreCase))
                    {
                        if (row.Cells[8].Value != null)
                        {
                            if (double.TryParse(row.Cells[8].Value.ToString(), out double valor))
                            {
                                row.Cells[8].Value = valor / 2;
                            }
                            else
                            {
                                MessageBox.Show("O valor na coluna 8 não é um número válido.", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            }
                        }
                    }
                }
            }
        }

        private void DividirValorColunaEPDMTabela2()
        {
            foreach (DataGridViewRow row in dataGridView2.Rows)
            {
                if (!row.IsNewRow)
                {
                    if (row.Cells[11].Value != null && row.Cells[11].Value.ToString().Equals("EPDM", StringComparison.OrdinalIgnoreCase))
                    {
                        if (row.Cells[8].Value != null)
                        {
                            if (double.TryParse(row.Cells[8].Value.ToString(), out double valor))
                            {
                                row.Cells[8].Value = valor / 2;
                            }
                            else
                            {
                                MessageBox.Show("O valor na coluna 8 não é um número válido.", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            }
                        }
                    }
                }
            }
        }
       
        public class ItemData
        {
            public string Nome { get; set; }
            public int Classe { get; set; }
            public string Norma { get; set; }
            public int Valor { get; set; }
        }

        private void ExtrairValores()
        {
            List<ItemData> listaValoresAntigos = new List<ItemData>();

            foreach (DataGridViewRow item in dataGridView2.Rows)
            {
                DataGridViewCell numerocell = item.Cells[8];
                DataGridViewCell nomecell = item.Cells[9];
                DataGridViewCell normacell = item.Cells[11];
                DataGridViewCell classecell = item.Cells[10];

                if (numerocell.Value != null && nomecell.Value.ToString().Contains("BM"))
                {
                    string nomeValue = nomecell.Value?.ToString() ?? string.Empty;
                    int numero = ExtrairNumero(nomeValue);

                    int classe = ExtrairClasse(classecell.Value?.ToString() ?? string.Empty);

                    var itemExistente = listaValoresAntigos.FirstOrDefault(x =>
                                        x.Nome == numero.ToString() &&
                                        x.Classe == classe &&
                                        x.Norma == normacell.Value.ToString());

                    if (itemExistente != null)
                    {
                        itemExistente.Valor += numerocell.Value != null ? Convert.ToInt32(numerocell.Value) : 0;
                    }
                    else
                    {
                        listaValoresAntigos.Add(new ItemData
                        {
                            Nome = numero.ToString(),
                            Classe = classe,
                            Norma = normacell.Value.ToString(),
                            Valor = numerocell.Value != null ? Convert.ToInt32(numerocell.Value) : 0
                        });
                    }
                }
            }

            foreach (DataGridViewRow item in dataGridView2.Rows)
            {
                DataGridViewCell currentcell = item.Cells[8];
                DataGridViewCell nomecell = item.Cells[9];

                if (currentcell.Value != null && (nomecell.Value.ToString().Contains("BM") || nomecell.Value.ToString().Contains("BUM")))
                {
                    double valorAtual = double.Parse(currentcell.Value.ToString());

                    // 0 a 10
                    if (valorAtual <= 10)
                    {
                        currentcell.Value = Math.Round(valorAtual + 1).ToString("0");
                    }
                    // Maior que 10 e até 250 (aumento de 5%)
                    else if (valorAtual > 10 && valorAtual <= 250)
                    {
                        currentcell.Value = Math.Round(valorAtual * 1.05).ToString("0");
                    }
                    // Maior que 250 e até 1000 (aumento de 2,5%)
                    else if (valorAtual > 250 && valorAtual <= 1000)
                    {
                        currentcell.Value = Math.Round(valorAtual * 1.025).ToString("0");
                    }
                    // Maior que 1000 e até 10000 (aumento de 2%)
                    else if (valorAtual > 1000 && valorAtual <= 10000)
                    {
                        currentcell.Value = Math.Round(valorAtual * 1.02).ToString("0");
                    }
                    // Maior que 10000 (aumento de 1%)
                    else
                    {
                        currentcell.Value = Math.Round(valorAtual * 1.01).ToString("0");
                    }
                }
            }

            List<ItemData> listaValoresNovos = new List<ItemData>();

            foreach (DataGridViewRow item in dataGridView2.Rows)
            {
                DataGridViewCell numerocell = item.Cells[8];
                DataGridViewCell nomecell = item.Cells[9];
                DataGridViewCell normacell = item.Cells[11];
                DataGridViewCell classecell = item.Cells[10];

                if (numerocell.Value != null && nomecell.Value.ToString().Contains("BM"))
                {
                    string nomeValue = nomecell.Value?.ToString() ?? string.Empty;
                    int numero = ExtrairNumero(nomeValue);

                    int classe = ExtrairClasse(classecell.Value?.ToString() ?? string.Empty);

                    var itemExistente = listaValoresNovos.FirstOrDefault(x =>
                                       x.Nome == numero.ToString() &&
                                       x.Classe == classe &&
                                       x.Norma == normacell.Value.ToString());

                    if (itemExistente != null)
                    {
                        itemExistente.Valor += numerocell.Value != null ? Convert.ToInt32(numerocell.Value) : 0;
                    }
                    else
                    {
                        listaValoresNovos.Add(new ItemData
                        {
                            Nome = numero.ToString(),
                            Classe = classe,
                            Norma = normacell.Value.ToString(),
                            Valor = numerocell.Value != null ? Convert.ToInt32(numerocell.Value) : 0 // Armazena o valor numérico
                        });
                    }
                }
            }

            List<ItemData> listaDiferencas = new List<ItemData>();

            foreach (var antigo in listaValoresAntigos)
            {
                var novo = listaValoresNovos.FirstOrDefault(x =>
                    x.Nome == antigo.Nome &&
                    x.Classe == antigo.Classe &&
                    x.Norma == antigo.Norma);

                if (novo != null)
                {
                    int diferenca = novo.Valor - antigo.Valor;

                    listaDiferencas.Add(new ItemData
                    {
                        Nome = antigo.Nome,
                        Classe = antigo.Classe,
                        Norma = antigo.Norma,
                        Valor = diferenca
                    });
                }
            }

            List<ItemData> listaDiferencasPorcas = new List<ItemData>();

            foreach (var antigo in listaValoresAntigos)
            {
                var novo = listaValoresNovos.FirstOrDefault(x =>
                    x.Nome == antigo.Nome &&
                    x.Classe == antigo.Classe &&
                    x.Norma == antigo.Norma);

                if (novo != null)
                {
                    int diferenca = novo.Valor - antigo.Valor;
                    string normaFinal = antigo.Norma.Contains("MAMA") ? "DIN 1587" : "ISO 4032";

                    listaDiferencasPorcas.Add(new ItemData
                    {
                        Nome = antigo.Nome,
                        Classe = antigo.Classe,
                        Norma = normaFinal,
                        Valor = diferenca
                    });
                }
            }

            foreach (var diferenca in listaDiferencasPorcas)
            {
                foreach (DataGridViewRow row in dataGridView2.Rows)
                {
                    DataGridViewCell numerocell = row.Cells[8];
                    DataGridViewCell nomecell = row.Cells[9];
                    DataGridViewCell normacell = row.Cells[11];
                    DataGridViewCell classecell = row.Cells[10];

                    if (numerocell.Value != null && nomecell.Value != null && nomecell.Value.ToString().Contains("NM"))
                    {
                        string nomeValue = nomecell.Value.ToString();
                        int numero = ExtrairNumero(nomeValue);
                        int classe = ExtrairClasse(classecell.Value?.ToString() ?? string.Empty);

                        if (numero.ToString() == diferenca.Nome &&
                            classe == diferenca.Classe &&
                            (normacell.Value.ToString() == diferenca.Norma))
                        {
                            if (numerocell.Value != null)
                            {
                                int valorAtual = Convert.ToInt32(numerocell.Value);
                                numerocell.Value = valorAtual + diferenca.Valor;
                            }

                            string mensagem = $"Porcas: Adicionada a diferença de {diferenca.Valor} para Nome: {diferenca.Nome}, Classe: {diferenca.Classe}, Norma: {diferenca.Norma}.";
                            break;
                        }
                        else
                        {
                            string debugMensagem = $"Porcas: Não corresponde: Nome: {nomeValue}, Classe: {classe}, Norma: {normacell.Value}. Esperado: Nome: {diferenca.Nome}, Classe: {diferenca.Classe}, Norma: {diferenca.Norma}.";
                        }
                    }
                }
            }

            List<ItemData> listaDiferencasAnilhas = new List<ItemData>();

            foreach (var antigo in listaValoresAntigos)
            {
                var novo = listaValoresNovos.FirstOrDefault(x =>
                    x.Nome == antigo.Nome &&
                    x.Classe == antigo.Classe &&
                    x.Norma == antigo.Norma);

                if (novo != null)
                {
                    int diferenca = novo.Valor - antigo.Valor;

                    string normaFinal;
                    if (antigo.Norma.Contains("MOLA"))
                    {
                        normaFinal = "DIN 127";
                    }
                    else if (antigo.Norma.Contains("UPN"))
                    {
                        normaFinal = "DIN 434";
                    }
                    else if (antigo.Norma.Contains("IPN"))
                    {
                        normaFinal = "DIN 435";
                    }
                    else
                    {
                        normaFinal = "ISO 7089";
                    }

                    listaDiferencasAnilhas.Add(new ItemData
                    {
                        Nome = antigo.Nome,
                        Classe = antigo.Classe,
                        Norma = normaFinal,
                        Valor = diferenca
                    });
                }

            }

            List<ItemData> listaDiferencasAnilhasENISO7089 = new List<ItemData>();

            foreach (var antigo in listaValoresAntigos)
            {
                var novo = listaValoresNovos.FirstOrDefault(x =>
                    x.Nome == antigo.Nome &&
                    x.Classe == antigo.Classe &&
                    x.Norma == antigo.Norma);

                if (novo != null)
                {
                    int diferenca = novo.Valor - antigo.Valor;

                    listaDiferencasAnilhasENISO7089.Add(new ItemData
                    {
                        Nome = antigo.Nome,
                        Classe = antigo.Classe,
                        Norma = "ISO 7089",
                        Valor = diferenca
                    });
                }

            }

            foreach (var diferenca in listaDiferencasAnilhas)
            {
                foreach (DataGridViewRow row in dataGridView2.Rows)
                {
                    DataGridViewCell numerocell = row.Cells[8];
                    DataGridViewCell nomecell = row.Cells[9];
                    DataGridViewCell normacell = row.Cells[11];
                    DataGridViewCell classecell = row.Cells[10];

                    if (numerocell.Value != null && nomecell.Value != null && nomecell.Value.ToString().Contains("WM"))
                    {
                        string nomeValue = nomecell.Value.ToString();
                        int numero = ExtrairNumero(nomeValue);

                        int classe = ExtrairClasse(classecell.Value?.ToString() ?? string.Empty);
                        int classediferente = 8;

                        if (classe == 300)
                        {
                            classediferente = 10;
                        }
                        else if (classe == 200)
                        {
                            classediferente = 8;

                        }
                        else if (classe == 140)
                        {
                            classediferente = 6;

                        }

                        if (numero.ToString() == diferenca.Nome &&
                            classediferente == diferenca.Classe &&
                            (normacell.Value.ToString() == diferenca.Norma))
                        {
                            if (numerocell.Value != null)
                            {
                                int valorAtual = Convert.ToInt32(numerocell.Value);
                                numerocell.Value = valorAtual + diferenca.Valor;
                            }

                            string mensagem = $"Anilhas: Adicionada a diferença de {diferenca.Valor} para Nome: {diferenca.Nome}, Classe: {diferenca.Classe}, Norma: {diferenca.Norma}.";
                            break;
                        }
                        else
                        {
                            string debugMensagem = $"Anilhas: Não corresponde: Nome: {numero}, Classe: {classediferente}, Norma: {normacell.Value}. Esperado: Nome: {diferenca.Nome}, Classe: {diferenca.Classe}, Norma: {diferenca.Norma}.";
                        }
                    }
                }
            }

            foreach (var diferenca in listaDiferencasAnilhasENISO7089)
            {
                foreach (DataGridViewRow row in dataGridView2.Rows)
                {
                    DataGridViewCell numerocell = row.Cells[8];
                    DataGridViewCell nomecell = row.Cells[9];
                    DataGridViewCell normacell = row.Cells[11];
                    DataGridViewCell classecell = row.Cells[10];

                    if (numerocell.Value != null && nomecell.Value != null && nomecell.Value.ToString().Contains("WM"))
                    {
                        string nomeValue = nomecell.Value.ToString();
                        int numero = ExtrairNumero(nomeValue);

                        int classe = ExtrairClasse(classecell.Value?.ToString() ?? string.Empty);
                        int classediferente = 8;

                        if (classe == 300)
                        {
                            classediferente = 10;
                        }
                        else if (classe == 200)
                        {
                            classediferente = 8;

                        }
                        else if (classe == 140)
                        {
                            classediferente = 6;

                        }

                        if (numero.ToString() == diferenca.Nome &&
                            classediferente == diferenca.Classe &&
                            (normacell.Value.ToString() == diferenca.Norma))
                        {
                            if (numerocell.Value != null)
                            {
                                int valorAtual = Convert.ToInt32(numerocell.Value);
                                numerocell.Value = valorAtual + diferenca.Valor;
                            }

                            string mensagem = $"Anilhas Normais : Adicionada a diferença de {diferenca.Valor} para Nome: {diferenca.Nome}, Classe: {diferenca.Classe}, Norma: {diferenca.Norma}.";
                            //MessageBox.Show(mensagem, "Diferença Adicionada");

                            break;
                        }
                        else
                        {
                            string debugMensagem = $"Anilhas Normais : Não corresponde: Nome: {numero}, Classe: {classediferente}, Norma: {normacell.Value}. Esperado: Nome: {diferenca.Nome}, Classe: {diferenca.Classe}, Norma: {diferenca.Norma}.";
                            //MessageBox.Show(debugMensagem);
                        }
                    }
                }
            }         
        }

        private int ExtrairNumero(string valor)
        {
            var match = System.Text.RegularExpressions.Regex.Match(valor, @"\d+");
            if (match.Success)
            {
                return Convert.ToInt32(match.Value);
            }
            return 0;
        }

        private int ExtrairClasse(string valor)
        {
            var match = System.Text.RegularExpressions.Regex.Match(valor, @"\d+");
            if (match.Success)
            {
                return Convert.ToInt32(match.Value);
            }
            return 0;
        }

        private void button9_Click(object sender, EventArgs e)
        {
            if (!Directory.Exists(Frm_Inico.CaminhoModelo + @"\listas"))
            {
                Directory.CreateDirectory(Frm_Inico.CaminhoModelo + @"\listas");
            }
            string Save = null;
            string Savesemlote = null;

            string numeroObra = lbl_numeroobra.Text.Trim();
            string fase = formpai.fase1000.Trim();

            if (Directory.Exists(Frm_Inico.PastaReservatorioFicheiros + fase))
            {
                Save = Frm_Inico.PastaReservatorioFicheiros + fase + "\\" + numeroObra + "F" + fase + "ComLote.csv";
                Savesemlote = Frm_Inico.PastaReservatorioFicheiros + fase + "\\" + numeroObra + "F" + fase + ".csv";

            }
            else
            {
                Directory.CreateDirectory(Frm_Inico.PastaReservatorioFicheiros + fase);
                Save = Frm_Inico.PastaReservatorioFicheiros + fase + "\\" + numeroObra + "F" + fase + "ComLote.csv";
                Savesemlote = Frm_Inico.PastaReservatorioFicheiros + fase + "\\" + numeroObra + "F" + fase + ".csv";
            }

            string pastaPartilhada = Frm_Inico.PastaPartilhada + "\\" + Frm_Inico.ano + "\\ARM\\" + numeroObra + "\\" + fase;

            if (!Directory.Exists(pastaPartilhada))
            {
                Directory.CreateDirectory(pastaPartilhada);
            }

            if (!Directory.Exists(Frm_Inico.PastaReservatorioFicheiros + fase + "\\20001"))
            {
                Directory.CreateDirectory(Frm_Inico.PastaReservatorioFicheiros + fase + "\\20001");
            }
            if (!Directory.Exists(Frm_Inico.PastaReservatorioFicheiros + fase + "\\20009"))
            {
                Directory.CreateDirectory(Frm_Inico.PastaReservatorioFicheiros + fase + "\\20009");
            }

            SaveToCSV(dataGridView1, Save);
            SaveToCSV2(dataGridView2, Savesemlote);

            Tekla.Structures.Model.Model m = new Tekla.Structures.Model.Model();
            string numeroProjeto = m.GetProjectInfo().ProjectNumber.Trim();
            MessageBox.Show("Lista exportada " + Environment.NewLine + " Fase:" + fase, "Exportação", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void guna2Button1_Click(object sender, EventArgs e)
        {
            for (int a = 0; a < dataGridView2.Rows.Count - 1; a++)
            {
                dataGridView2.Rows[a].Cells[18].Value = dateTimePicker1.Text;
            }
        }
              
        private void MudarColunaClasseTabela1()
        {
            Dictionary<string, string> mapeamentoClasseBMeVrsm = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            {
                    { "8.8 ZINCADO", "8.8 ZINCADO" },
                    { "8.8-SB ZINCADO", "8.8 ZINCADO" },
                    { "10.9-HV ZINCADO", "10.9 ZINCADO" },
                    { "8,8 ZINCADO 2,1", "8.8 ZINCADO" },
                    { "8,8 ZINCADO 2,2", "8.8 ZINCADO" },
                    { "8,8 ZINCADO 3,1", "8.8 ZINCADO" },
                    { "8,8 ZINCADO 3,2", "8.8 ZINCADO" },
                    { "10,9 ZINCADO 2,1", "10.9 ZINCADO" },
                    { "10,9 ZINCADO 2,2", "10.9 ZINCADO" },
                    { "10,9 ZINCADO 3,1", "10.9 ZINCADO" },
                    { "10,9 ZINCADO 3,2", "10.9 ZINCADO" },

                    { "8.8-SB PRETO", "8.8 PRETO" },
                    { "10.9-HV PRETO", "10.9 PRETO" },
                    { "8,8 PRETO 2,1", "8.8 PRETO" },
                    { "8,8 PRETO 2,2", "8.8 PRETO" },
                    { "8,8 PRETO 3,1", "8.8 PRETO" },
                    { "8,8 PRETO 3,2", "8.8 PRETO" },
                    { "10,9 PRETO 2,1", "10.9 PRETO" },
                    { "10,9 PRETO 2,2", "10.9 PRETO" },
                    { "10,9 PRETO 3,1", "10.9 PRETO" },
                    { "10,9 PRETO 3,2", "10.9 PRETO" },

                    { "8.8-SB GALVANIZADO", "8.8 GALVANIZADO" },
                    { "10.9-HV GALVANIZADO", "10.9 GALVANIZADO" },
                    { "8,8 GALVANIZADO 2,1", "8.8 GALVANIZADO" },
                    { "8,8 GALVANIZADO 2,2", "8.8 GALVANIZADO" },
                    { "8,8 GALVANIZADO 3,1", "8.8 GALVANIZADO" },
                    { "8,8 GALVANIZADO 3,2", "8.8 GALVANIZADO" },
                    { "10,9 GALVANIZADO 2,1", "10.9 GALVANIZADO" },
                    { "10,9 GALVANIZADO 2,2", "10.9 GALVANIZADO" },
                    { "10,9 GALVANIZADO 3,1", "10.9 GALVANIZADO" },
                    { "10,9 GALVANIZADO 3,2", "10.9 GALVANIZADO" }
            };
            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                if (row.IsNewRow) continue;

                var perfil = row.Cells[9].Value?.ToString();
                var classeOriginal = row.Cells[10].Value?.ToString();

                if ((perfil?.StartsWith("BM", StringComparison.OrdinalIgnoreCase) ?? false) ||
                   (perfil?.StartsWith("VRSM", StringComparison.OrdinalIgnoreCase) ?? false))
                {
                    if (mapeamentoClasseBMeVrsm.TryGetValue(classeOriginal, out string novaClasse))
                    {
                        row.Cells[10].Value = novaClasse;
                    }
                }
            }

            Dictionary<string, string> mapeamentoClasseNW = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            {
                    { "8.8 ZINCADO", "8 ZINCADO" },
                    { "8.8-SB ZINCADO", "8 ZINCADO" },
                    { "10.9-HV ZINCADO", "10 ZINCADO" },
                    { "8,8 ZINCADO 2,1", "8 ZINCADO" },
                    { "8,8 ZINCADO 2,2", "8 ZINCADO" },
                    { "8,8 ZINCADO 3,1", "8 ZINCADO" },
                    { "8,8 ZINCADO 3,2", "8 ZINCADO" },
                    { "10,9 ZINCADO 2,1", "10 ZINCADO" },
                    { "10,9 ZINCADO 2,2", "10 ZINCADO" },
                    { "10,9 ZINCADO 3,1", "10 ZINCADO" },
                    { "10,9 ZINCADO 3,2", "10 ZINCADO" },

                    { "8.8-SB PRETO", "8 PRETO" },
                    { "10.9-HV PRETO", "10 PRETO" },
                    { "8,8 PRETO 2,1", "8 PRETO" },
                    { "8,8 PRETO 2,2", "8 PRETO" },
                    { "8,8 PRETO 3,1", "8 PRETO" },
                    { "8,8 PRETO 3,2", "8 PRETO" },
                    { "10,9 PRETO 2,1", "10 PRETO" },
                    { "10,9 PRETO 2,2", "10 PRETO" },
                    { "10,9 PRETO 3,1", "10 PRETO" },
                    { "10,9 PRETO 3,2", "10 PRETO" },

                    { "8.8-SB GALVANIZADO", "8 GALVANIZADO" },
                    { "10.9-HV GALVANIZADO", "10 GALVANIZADO" },
                    { "8,8 GALVANIZADO 2,1", "8 GALVANIZADO" },
                    { "8,8 GALVANIZADO 2,2", "8 GALVANIZADO" },
                    { "8,8 GALVANIZADO 3,1", "8 GALVANIZADO" },
                    { "8,8 GALVANIZADO 3,2", "8 GALVANIZADO" },
                    { "10,9 GALVANIZADO 2,1", "10 GALVANIZADO" },
                    { "10,9 GALVANIZADO 2,2", "10 GALVANIZADO" },
                    { "10,9 GALVANIZADO 3,1", "10 GALVANIZADO" },
                    { "10,9 GALVANIZADO 3,2", "10 GALVANIZADO" }
            };

            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                if (row.IsNewRow) continue;

                var perfil = row.Cells[9].Value?.ToString();
                var classeOriginal = row.Cells[10].Value?.ToString();

                if ((perfil?.StartsWith("NM", StringComparison.OrdinalIgnoreCase) ?? false))
                {
                    if (mapeamentoClasseNW.TryGetValue(classeOriginal, out string novaClasse))
                    {
                        row.Cells[10].Value = novaClasse;
                    }
                }
            }

            Dictionary<string, string> mapeamentoClasseWM = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            {
                    { "8.8 ZINCADO", "200 HV ZINCADO" },
                    { "8.8-SB ZINCADO", "200 HV ZINCADO" },
                    { "10.9-HV ZINCADO", "300 HV ZINCADO" },
                    { "8,8 ZINCADO 2,1", "200 HV ZINCADO" },
                    { "8,8 ZINCADO 2,2", "200 HV ZINCADO" },
                    { "8,8 ZINCADO 3,1", "200 HV ZINCADO" },
                    { "8,8 ZINCADO 3,2", "200 HV ZINCADO" },
                    { "10,9 ZINCADO 2,1", "300 HV ZINCADO" },
                    { "10,9 ZINCADO 2,2", "300 HV ZINCADO" },
                    { "10,9 ZINCADO 3,1", "300 HV ZINCADO" },
                    { "10,9 ZINCADO 3,2", "300 HV ZINCADO" },

                    { "8.8-SB PRETO", "200 HV PRETO" },
                    { "10.9-HV PRETO", "300 HV PRETO" },
                    { "8,8 PRETO 2,1", "200 HV PRETO" },
                    { "8,8 PRETO 2,2", "200 HV PRETO" },
                    { "8,8 PRETO 3,1", "200 HV PRETO" },
                    { "8,8 PRETO 3,2", "200 HV PRETO" },
                    { "10,9 PRETO 2,1", "300 HV PRETO" },
                    { "10,9 PRETO 2,2", "300 HV PRETO" },
                    { "10,9 PRETO 3,1", "300 HV PRETO" },
                    { "10,9 PRETO 3,2", "300 HV PRETO" },

                    { "8.8-SB GALVANIZADO", "200 HV GALVANIZADO" },
                    { "10.9-HV GALVANIZADO", "300 HV GALVANIZADO" },
                    { "8,8 GALVANIZADO 2,1", "200 HV GALVANIZADO" },
                    { "8,8 GALVANIZADO 2,2", "200 HV GALVANIZADO" },
                    { "8,8 GALVANIZADO 3,1", "200 HV GALVANIZADO" },
                    { "8,8 GALVANIZADO 3,2", "200 HV GALVANIZADO" },
                    { "10,9 GALVANIZADO 2,1", "300 HV GALVANIZADO" },
                    { "10,9 GALVANIZADO 2,2", "300 HV GALVANIZADO" },
                    { "10,9 GALVANIZADO 3,1", "300 HV GALVANIZADO" },
                    { "10,9 GALVANIZADO 3,2", "300 HV GALVANIZADO" }
            };
            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                if (row.IsNewRow) continue;

                var perfil = row.Cells[9].Value?.ToString();
                var classeOriginal = row.Cells[10].Value?.ToString();

                if (perfil.StartsWith("WM", StringComparison.OrdinalIgnoreCase))
                {
                    if (mapeamentoClasseWM.TryGetValue(classeOriginal, out string novaClasse))
                    {
                        row.Cells[10].Value = novaClasse;
                    }
                }
            }
        }

        private void MudarColunaClasseTabela2()
        {
            Dictionary<string, string> mapeamentoClasseBMeVrsm = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            {
                    { "8.8 ZINCADO", "8.8 ZINCADO" },
                    { "8.8-SB ZINCADO", "8.8 ZINCADO" },
                    { "10.9-HV ZINCADO", "10.9 ZINCADO" },
                    { "8,8 ZINCADO 2,1", "8.8 ZINCADO" },
                    { "8,8 ZINCADO 2,2", "8.8 ZINCADO" },
                    { "8,8 ZINCADO 3,1", "8.8 ZINCADO" },
                    { "8,8 ZINCADO 3,2", "8.8 ZINCADO" },
                    { "10,9 ZINCADO 2,1", "10.9 ZINCADO" },
                    { "10,9 ZINCADO 2,2", "10.9 ZINCADO" },
                    { "10,9 ZINCADO 3,1", "10.9 ZINCADO" },
                    { "10,9 ZINCADO 3,2", "10.9 ZINCADO" },

                    { "8.8-SB PRETO", "8.8 PRETO" },
                    { "10.9-HV PRETO", "10.9 PRETO" },
                    { "8,8 PRETO 2,1", "8.8 PRETO" },
                    { "8,8 PRETO 2,2", "8.8 PRETO" },
                    { "8,8 PRETO 3,1", "8.8 PRETO" },
                    { "8,8 PRETO 3,2", "8.8 PRETO" },
                    { "10,9 PRETO 2,1", "10.9 PRETO" },
                    { "10,9 PRETO 2,2", "10.9 PRETO" },
                    { "10,9 PRETO 3,1", "10.9 PRETO" },
                    { "10,9 PRETO 3,2", "10.9 PRETO" },

                    { "8.8-SB GALVANIZADO", "8.8 GALVANIZADO" },
                    { "10.9-HV GALVANIZADO", "10.9 GALVANIZADO" },
                    { "8,8 GALVANIZADO 2,1", "8.8 GALVANIZADO" },
                    { "8,8 GALVANIZADO 2,2", "8.8 GALVANIZADO" },
                    { "8,8 GALVANIZADO 3,1", "8.8 GALVANIZADO" },
                    { "8,8 GALVANIZADO 3,2", "8.8 GALVANIZADO" },
                    { "10,9 GALVANIZADO 2,1", "10.9 GALVANIZADO" },
                    { "10,9 GALVANIZADO 2,2", "10.9 GALVANIZADO" },
                    { "10,9 GALVANIZADO 3,1", "10.9 GALVANIZADO" },
                    { "10,9 GALVANIZADO 3,2", "10.9 GALVANIZADO" }
            };
            foreach (DataGridViewRow row in dataGridView2.Rows)
            {
                if (row.IsNewRow) continue;

                var perfil = row.Cells[9].Value?.ToString();
                var classeOriginal = row.Cells[10].Value?.ToString();

                if ((perfil?.StartsWith("BM", StringComparison.OrdinalIgnoreCase) ?? false) ||
                   (perfil?.StartsWith("VRSM", StringComparison.OrdinalIgnoreCase) ?? false))
                {
                    if (mapeamentoClasseBMeVrsm.TryGetValue(classeOriginal, out string novaClasse))
                    {
                        row.Cells[10].Value = novaClasse;
                    }
                }
            }

            Dictionary<string, string> mapeamentoClasseNW = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            {
                    { "8.8 ZINCADO", "8 ZINCADO" },
                    { "8.8-SB ZINCADO", "8 ZINCADO" },
                    { "10.9-HV ZINCADO", "10 ZINCADO" },
                    { "8,8 ZINCADO 2,1", "8 ZINCADO" },
                    { "8,8 ZINCADO 2,2", "8 ZINCADO" },
                    { "8,8 ZINCADO 3,1", "8 ZINCADO" },
                    { "8,8 ZINCADO 3,2", "8 ZINCADO" },
                    { "10,9 ZINCADO 2,1", "10 ZINCADO" },
                    { "10,9 ZINCADO 2,2", "10 ZINCADO" },
                    { "10,9 ZINCADO 3,1", "10 ZINCADO" },
                    { "10,9 ZINCADO 3,2", "10 ZINCADO" },

                    { "8.8-SB PRETO", "8 PRETO" },
                    { "10.9-HV PRETO", "10 PRETO" },
                    { "8,8 PRETO 2,1", "8 PRETO" },
                    { "8,8 PRETO 2,2", "8 PRETO" },
                    { "8,8 PRETO 3,1", "8 PRETO" },
                    { "8,8 PRETO 3,2", "8 PRETO" },
                    { "10,9 PRETO 2,1", "10 PRETO" },
                    { "10,9 PRETO 2,2", "10 PRETO" },
                    { "10,9 PRETO 3,1", "10 PRETO" },
                    { "10,9 PRETO 3,2", "10 PRETO" },

                    { "8.8-SB GALVANIZADO", "8 GALVANIZADO" },
                    { "10.9-HV GALVANIZADO", "10 GALVANIZADO" },
                    { "8,8 GALVANIZADO 2,1", "8 GALVANIZADO" },
                    { "8,8 GALVANIZADO 2,2", "8 GALVANIZADO" },
                    { "8,8 GALVANIZADO 3,1", "8 GALVANIZADO" },
                    { "8,8 GALVANIZADO 3,2", "8 GALVANIZADO" },
                    { "10,9 GALVANIZADO 2,1", "10 GALVANIZADO" },
                    { "10,9 GALVANIZADO 2,2", "10 GALVANIZADO" },
                    { "10,9 GALVANIZADO 3,1", "10 GALVANIZADO" },
                    { "10,9 GALVANIZADO 3,2", "10 GALVANIZADO" }
            };

            foreach (DataGridViewRow row in dataGridView2.Rows)
            {
                if (row.IsNewRow) continue;

                var perfil = row.Cells[9].Value?.ToString();
                var classeOriginal = row.Cells[10].Value?.ToString();

                if ((perfil?.StartsWith("NM", StringComparison.OrdinalIgnoreCase) ?? false))
                {
                    if (mapeamentoClasseNW.TryGetValue(classeOriginal, out string novaClasse))
                    {
                        row.Cells[10].Value = novaClasse;
                    }
                }
            }

            Dictionary<string, string> mapeamentoClasseWM = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            {
                    { "8.8 ZINCADO", "200 HV ZINCADO" },
                    { "8.8-SB ZINCADO", "200 HV ZINCADO" },
                    { "10.9-HV ZINCADO", "300 HV ZINCADO" },
                    { "8,8 ZINCADO 2,1", "200 HV ZINCADO" },
                    { "8,8 ZINCADO 2,2", "200 HV ZINCADO" },
                    { "8,8 ZINCADO 3,1", "200 HV ZINCADO" },
                    { "8,8 ZINCADO 3,2", "200 HV ZINCADO" },
                    { "10,9 ZINCADO 2,1", "300 HV ZINCADO" },
                    { "10,9 ZINCADO 2,2", "300 HV ZINCADO" },
                    { "10,9 ZINCADO 3,1", "300 HV ZINCADO" },
                    { "10,9 ZINCADO 3,2", "300 HV ZINCADO" },

                    { "8.8-SB PRETO", "200 HV PRETO" },
                    { "10.9-HV PRETO", "300 HV PRETO" },
                    { "8,8 PRETO 2,1", "200 HV PRETO" },
                    { "8,8 PRETO 2,2", "200 HV PRETO" },
                    { "8,8 PRETO 3,1", "200 HV PRETO" },
                    { "8,8 PRETO 3,2", "200 HV PRETO" },
                    { "10,9 PRETO 2,1", "300 HV PRETO" },
                    { "10,9 PRETO 2,2", "300 HV PRETO" },
                    { "10,9 PRETO 3,1", "300 HV PRETO" },
                    { "10,9 PRETO 3,2", "300 HV PRETO" },

                    { "8.8-SB GALVANIZADO", "200 HV GALVANIZADO" },
                    { "10.9-HV GALVANIZADO", "300 HV GALVANIZADO" },
                    { "8,8 GALVANIZADO 2,1", "200 HV GALVANIZADO" },
                    { "8,8 GALVANIZADO 2,2", "200 HV GALVANIZADO" },
                    { "8,8 GALVANIZADO 3,1", "200 HV GALVANIZADO" },
                    { "8,8 GALVANIZADO 3,2", "200 HV GALVANIZADO" },
                    { "10,9 GALVANIZADO 2,1", "300 HV GALVANIZADO" },
                    { "10,9 GALVANIZADO 2,2", "300 HV GALVANIZADO" },
                    { "10,9 GALVANIZADO 3,1", "300 HV GALVANIZADO" },
                    { "10,9 GALVANIZADO 3,2", "300 HV GALVANIZADO" }
            };
            foreach (DataGridViewRow row in dataGridView2.Rows)
            {
                if (row.IsNewRow) continue;

                var perfil = row.Cells[9].Value?.ToString();
                var classeOriginal = row.Cells[10].Value?.ToString();

                if (perfil.StartsWith("WM", StringComparison.OrdinalIgnoreCase))
                {
                    if (mapeamentoClasseWM.TryGetValue(classeOriginal, out string novaClasse))
                    {
                        row.Cells[10].Value = novaClasse;
                    }
                }
            }
        }

        private void EmcasodeerrodeReqEspecialTabela1()
        {
            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                if (row.IsNewRow) continue;

                var profile = row.Cells[9].Value?.ToString();
                var valorAtual = row.Cells[11].Value?.ToString();

                if (string.IsNullOrEmpty(valorAtual) || valorAtual == "0")
                {
                    string resultado;

                    if (profile != null && profile.StartsWith("VRSM", StringComparison.OrdinalIgnoreCase))
                    {
                        resultado = "DIN976";
                    }
                    else if (profile != null && profile.StartsWith("NM", StringComparison.OrdinalIgnoreCase))
                    {
                        resultado = "ISO 4032";
                    }
                    else if (profile != null && profile.StartsWith("WM", StringComparison.OrdinalIgnoreCase))
                    {
                        resultado = "ISO 7089";
                    }
                    else if (profile != null && profile.StartsWith("BM", StringComparison.OrdinalIgnoreCase))
                    {
                        resultado = "ISO 4017";
                    }
                    else
                    {
                        resultado = "";
                    }

                    row.Cells[11].Value = resultado;
                }
            }
        }

        private void EmcasodeerrodeReqEspecialTabela2()
        {
            foreach (DataGridViewRow row in dataGridView2.Rows)
            {
                if (row.IsNewRow) continue;

                var profile = row.Cells[9].Value?.ToString();
                var valorAtual = row.Cells[11].Value?.ToString();

                if (string.IsNullOrEmpty(valorAtual) || valorAtual == "0")
                {
                    string resultado;

                    if (profile != null && profile.StartsWith("VRSM", StringComparison.OrdinalIgnoreCase))
                    {
                        resultado = "DIN976";
                    }
                    else if (profile != null && profile.StartsWith("NM", StringComparison.OrdinalIgnoreCase))
                    {
                        resultado = "ISO 4032";
                    }
                    else if (profile != null && profile.StartsWith("WM", StringComparison.OrdinalIgnoreCase))
                    {
                        resultado = "ISO 7089";
                    }
                    else if (profile != null && profile.StartsWith("BM", StringComparison.OrdinalIgnoreCase))
                    {
                        resultado = "ISO 4017";
                    }
                    else
                    {
                        resultado = "";
                    }

                    row.Cells[11].Value = resultado;
                }
            }
        }

        private void SubstituirPorcasEAnilhasSoltas1()
        {
            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                if (row.IsNewRow) continue;

                var valorCelula = row.Cells[11].Value?.ToString();

                if (!string.IsNullOrEmpty(valorCelula))
                {
                    string valorLower = valorCelula.ToLower();

                    if (valorLower.Contains("porcas soltas"))
                    {
                        row.Cells[11].Value = "ISO 4032";
                    }
                    else if (valorLower.Contains("anilhas soltas"))
                    {
                        row.Cells[11].Value = "ISO 7089";
                    }
                }
            }
        }

        private void SubstituirPorcasEAnilhasSoltas2()
        {
            foreach (DataGridViewRow row in dataGridView2.Rows)
            {
                if (row.IsNewRow) continue;

                var valorCelula = row.Cells[11].Value?.ToString();

                if (!string.IsNullOrEmpty(valorCelula))
                {
                    string valorLower = valorCelula.ToLower();

                    if (valorLower.Contains("porcas soltas"))
                    {
                        row.Cells[11].Value = "ISO 4032";
                    }
                    else if (valorLower.Contains("anilhas soltas"))
                    {
                        row.Cells[11].Value = "ISO 7089";
                    }
                }
            }
        }

        private void ConsolidarLinhasDuplicadas1()
        {
            var linhasUnicas = new Dictionary<string, DataGridViewRow>();
            var linhasParaRemover = new List<DataGridViewRow>();

            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                if (row.IsNewRow) continue;

                string perfil = row.Cells[9].Value?.ToString()?.Trim() ?? "";
                string classe = row.Cells[10].Value?.ToString()?.Trim() ?? "";
                string requisito = row.Cells[11].Value?.ToString()?.Trim() ?? "";

                string chave = $"{perfil}|{classe}|{requisito}";

                if (linhasUnicas.ContainsKey(chave))
                {
                    var linhaExistente = linhasUnicas[chave];
                    int quantidadeExistente = Convert.ToInt32(linhaExistente.Cells[8].Value ?? 0);
                    int quantidadeAtual = Convert.ToInt32(row.Cells[8].Value ?? 0);
                    linhaExistente.Cells[8].Value = quantidadeExistente + quantidadeAtual;
                    linhasParaRemover.Add(row);
                }
                else
                {
                    linhasUnicas[chave] = row;
                }
            }

            foreach (var linha in linhasParaRemover)
            {
                dataGridView1.Rows.Remove(linha);
            }
        }

        private void ConsolidarLinhasDuplicadas2()
        {
            var linhasUnicas = new Dictionary<string, DataGridViewRow>();
            var linhasParaRemover = new List<DataGridViewRow>();

            foreach (DataGridViewRow row in dataGridView2.Rows)
            {
                if (row.IsNewRow) continue;

                string perfil = row.Cells[9].Value?.ToString()?.Trim() ?? "";
                string classe = row.Cells[10].Value?.ToString()?.Trim() ?? "";
                string requisito = row.Cells[11].Value?.ToString()?.Trim() ?? "";

                string chave = $"{perfil}|{classe}|{requisito}";

                if (linhasUnicas.ContainsKey(chave))
                {
                    var linhaExistente = linhasUnicas[chave];
                    int quantidadeExistente = Convert.ToInt32(linhaExistente.Cells[8].Value ?? 0);
                    int quantidadeAtual = Convert.ToInt32(row.Cells[8].Value ?? 0);
                    linhaExistente.Cells[8].Value = quantidadeExistente + quantidadeAtual;
                    linhasParaRemover.Add(row);
                }
                else
                {
                    linhasUnicas[chave] = row;
                }
            }

            foreach (var linha in linhasParaRemover)
            {
                dataGridView2.Rows.Remove(linha);
            }
        }

        private void AjustarColunasDataGridView1()
        {
            List<DateTime> data = new List<DateTime>();

            for (int a = 0; a < dataGridView1.Rows.Count - 1; a++)
            {
                for (int b = 0; b < dataGridView1.ColumnCount - 1; b++)
                {
                    if (b == 0)
                    {
                        dataGridView1.Rows[a].Cells[0].Value = formpai.fase1000;
                    }
                    else if (b == 3)
                    {
                        dataGridView1.Rows[a].Cells[3].Value = "2." + lbl_numeroobra.Text + "." + dataGridView1.Rows[a].Cells[0].Value + "." + (a + 1);
                    }
                    else if (b == 4)
                    {
                        dataGridView1.Rows[a].Cells[4].Value = "2." + lbl_numeroobra.Text + "." + dataGridView1.Rows[a].Cells[1].Value + "." + formpai.fase1000 + "H" + (a + 1);
                    }
                    else if (b == 18)
                    {
                        try
                        {
                            dataGridView1.Rows[a].Cells[b].Value = dataGridView1.Rows[a].Cells[b].Value.ToString().Replace(".", "/").Replace("-", "/").Replace("_", "/").Trim();
                            data.Add(Convert.ToDateTime(dataGridView1.Rows[a].Cells[b].Value.ToString()));
                        }
                        catch (Exception)
                        {
                        }
                    }
                    else
                    {
                        dataGridView1.Rows[a].Cells[b].Value = dataGridView1.Rows[a].Cells[b].Value.ToString().Trim();
                    }
                }
            }

            dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
            dataGridView1.Columns[18].DefaultCellStyle.Format = "dd/MM/yyyy";
            data.Sort();
            try
            {
                dateTimePicker1.Text = data[0].ToShortDateString();
            }
            catch (Exception)
            {
            }
        }

        private void AjustarColunasDataGridView2()
        {
            List<DateTime> data = new List<DateTime>();

            for (int a = 0; a < dataGridView2.Rows.Count - 1; a++)
            {
                for (int b = 0; b < dataGridView2.ColumnCount - 1; b++)
                {
                    if (b == 0)
                    {
                        dataGridView2.Rows[a].Cells[0].Value = formpai.fase1000;
                    }
                    else if (b == 3)
                    {
                        dataGridView2.Rows[a].Cells[3].Value = "2." + lbl_numeroobra.Text + "." + dataGridView2.Rows[a].Cells[0].Value + "." + (a + 1);
                    }
                    else if (b == 4)
                    {
                        dataGridView2.Rows[a].Cells[4].Value = "2." + lbl_numeroobra.Text + "." + dataGridView2.Rows[a].Cells[1].Value + "." + formpai.fase1000 + "H" + (a + 1);
                    }
                    else if (b == 18)
                    {
                        try
                        {
                            dataGridView2.Rows[a].Cells[b].Value = dataGridView2.Rows[a].Cells[b].Value.ToString().Replace(".", "/").Replace("-", "/").Replace("_", "/").Trim();
                            data.Add(Convert.ToDateTime(dataGridView2.Rows[a].Cells[b].Value.ToString()));
                        }
                        catch (Exception)
                        {
                        }
                    }
                    else
                    {
                        dataGridView2.Rows[a].Cells[b].Value = dataGridView2.Rows[a].Cells[b].Value.ToString().Trim();
                    }
                }
            }

            dataGridView2.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
            dataGridView2.Columns[18].DefaultCellStyle.Format = "dd/MM/yyyy";
            data.Sort();
            try
            {
                dateTimePicker1.Text = data[0].ToShortDateString();
            }
            catch (Exception)
            {
            }
        }


        private void ButttonVerTabela1_Click(object sender, EventArgs e)
        {
            if (ButttonVerTabela1.Text == "Ver Tabela Real")
            {
                ButttonVerTabela1.Text = "Fechar Tabela Real";
                this.Size = new Size(1500, 1050);              
                dataGridView1.Visible = true;                
                this.MinimumSize = new Size(1500, 1050);
                this.MaximumSize = new Size(2700, 1050);
            }
            else
            {
                ButttonVerTabela1.Text = "Ver Tabela Real";
                this.Size = new Size(1500, 620);             
                dataGridView1.Visible = false;
                this.MinimumSize = new Size(1500, 620);
                this.MaximumSize = new Size(2700, 620);
            }           

        }

        private void guna2Button1_Click_1(object sender, EventArgs e)
        {
            dataGridView2.ReadOnly = !dataGridView2.ReadOnly;
        }
    }
}
