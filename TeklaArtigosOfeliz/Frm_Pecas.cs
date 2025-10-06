using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using static TeklaArtigosOfeliz.Frm_CriarFase;
using System.Text.RegularExpressions;
using Tekla.Structures.Model;
using System.Runtime.InteropServices;
using System.Threading;
using WindowsInput.Native;
using WindowsInput;
using System.Diagnostics;
using System.Web.UI;

namespace TeklaArtigosOfeliz
{
    public partial class Frm_Pecas : Form
    {
        Frm_ListaOFeliz _FrmPai;
        Frm_Inico _Formprin;

       
        public Frm_Pecas(Frm_ListaOFeliz pai, Frm_Inico formprin)
        {
            InitializeComponent();
            _FrmPai = pai;
            _Formprin = formprin;
            CarregaDados();
            this.FormClosed -= Frm_Pecas_FormClosed;
            this.FormClosed += Frm_Pecas_FormClosed;
        }

        void CarregaDados()
        {
            if (_FrmPai.listacarregadaexc3)
            {
                int ano = 0;
                if (_FrmPai.lbl_numeroobra.Text.ToLower().Contains("pt"))
                {
                    ano = int.Parse("20" + _FrmPai.lbl_numeroobra.Text.Substring(2, 2));
                }
                else
                {
                    ano = int.Parse("20" + _FrmPai.lbl_numeroobra.Text.Substring(0, 2));
                }
                Frm_Inico.ano = ano.ToString();


                Frm_Inico.PastaReservatorioFicheiros = @"\\Marconi\COMPANY SHARED FOLDER\OFELIZ\OFM\2.AN\2.CM\DP\1 Obras\" + ano + "\\" + _FrmPai.lbl_numeroobra.Text.Trim() + "\\1.9 Gestão de fabrico\\";

                for (int i = 0; i < _FrmPai.dataGridView1.Rows.Count - 1; i++)
                {
                    bool resultado = false;
                    string Pecas = _FrmPai.dataGridView1.Rows[i].Cells[4].Value.ToString().Split('.').Last().Replace(" ", "");

                    if (_FrmPai.dataGridView1.Rows[i].Cells[22].Value.ToString().ToLower().Contains("opção"))
                    {
                        resultado = ValidaExistente(Pecas, _FrmPai.dataGridView1.Rows[i].Cells[0].Value.ToString().Replace(" ", ""), _FrmPai.dataGridView1.Rows[i].Cells[1].Value.ToString().Replace(" ", ""));

                    }
                    else
                    {
                        resultado = ValidaExistente(_FrmPai.dataGridView1.Rows[i].Cells[4].Value.ToString().Split('.').Last().Replace(" ", ""), _FrmPai.dataGridView1.Rows[i].Cells[0].Value.ToString().Replace(" ", ""), _FrmPai.dataGridView1.Rows[i].Cells[1].Value.ToString().Replace(" ", ""));
                    }



                    //!(_FrmPai.dataGridView1.Rows[i].Cells[2].Value.ToString().ToLower().Contains("conjunto")) &&
                    if ((resultado == true) && !(_FrmPai.dataGridView1.Rows[i].Cells[22].Value.ToString().ToLower().Contains("opção 9")))
                    {
                        if (_FrmPai.dataGridView1.Rows[i].Cells[22].Value.ToString().ToLower().Contains("opção"))
                        {
                            dataGridView1.Rows.Add(_FrmPai.dataGridView1.Rows[i].Cells[0].Value.ToString().Replace(" ", ""), _FrmPai.dataGridView1.Rows[i].Cells[1].Value.ToString().Replace(" ", ""), Pecas, _FrmPai.dataGridView1.Rows[i].Cells[9].Value.ToString().ToString().Replace(" ", ""), _FrmPai.dataGridView1.Rows[i].Cells[10].Value.ToString().ToString().Replace(" ", ""), _FrmPai.dataGridView1.Rows[i].Cells[20].Value.ToString().Replace(" ", ""), _FrmPai.dataGridView1.Rows[i].Cells[22].Value);
                        }
                        else
                        {
                            dataGridView1.Rows.Add(_FrmPai.dataGridView1.Rows[i].Cells[0].Value.ToString().Replace(" ", ""), _FrmPai.dataGridView1.Rows[i].Cells[1].Value.ToString().Replace(" ", ""), _FrmPai.dataGridView1.Rows[i].Cells[4].Value.ToString().Split('.').Last().Replace(" ", ""), _FrmPai.dataGridView1.Rows[i].Cells[9].Value.ToString().ToString().Replace(" ", ""), _FrmPai.dataGridView1.Rows[i].Cells[10].Value.ToString().ToString().Replace(" ", ""), _FrmPai.dataGridView1.Rows[i].Cells[20].Value.ToString().Replace(" ", ""), _FrmPai.dataGridView1.Rows[i].Cells[22].Value);
                        }

                        dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;

                    }
                }
            }
            else
            {


                for (int i = 0; i < _FrmPai.dataGridView1.Rows.Count - 1; i++)
                {
                    bool resultado = false;
                    resultado = ValidaExistente(_FrmPai.dataGridView1.Rows[i].Cells[4].Value.ToString().Split('.').Last().Replace(" ", ""), _FrmPai.dataGridView1.Rows[i].Cells[0].Value.ToString().Replace(" ", ""), _FrmPai.dataGridView1.Rows[i].Cells[1].Value.ToString().Replace(" ", ""));

                    //!(_FrmPai.dataGridView1.Rows[i].Cells[2].Value.ToString().ToLower().Contains("conjunto")) &&
                    if ((resultado == true) && !(_FrmPai.dataGridView1.Rows[i].Cells[22].Value.ToString().ToLower().Contains("opção 9")))
                    {

                        dataGridView1.Rows.Add(_FrmPai.dataGridView1.Rows[i].Cells[0].Value.ToString().Replace(" ", ""), _FrmPai.dataGridView1.Rows[i].Cells[1].Value.ToString().Replace(" ", ""), _FrmPai.dataGridView1.Rows[i].Cells[4].Value.ToString().Split('.').Last().Replace(" ", ""), _FrmPai.dataGridView1.Rows[i].Cells[9].Value.ToString().ToString().Replace(" ", ""), _FrmPai.dataGridView1.Rows[i].Cells[10].Value.ToString().ToString().Replace(" ", ""), _FrmPai.dataGridView1.Rows[i].Cells[20].Value.ToString().Replace(" ", ""), _FrmPai.dataGridView1.Rows[i].Cells[22].Value);


                        dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;

                    }
                }
            }
        }

        bool ValidaExistente(string peca, string fase, string lote)
        {
            bool resultado = true;
            for (int i = 0; i < dataGridView1.Rows.Count - 1; i++)
            {
                if (dataGridView1.Rows[i].Cells[2].Value.ToString().Replace(" ", "") == peca && dataGridView1.Rows[i].Cells[0].Value.ToString().Replace(" ", "") == fase)
                {
                    resultado = false;
                }
            }
            return resultado;
        }

        private void FrmPecas_Load(object sender, EventArgs e)
        {

        }

        public Button Botao1 => button1; 


        private void button1_Click(object sender, EventArgs e)
        {
            bool PASTA = true;
            if (Directory.Exists(Frm_Inico.PastaReservatorioFicheiros + int.Parse(dataGridView1.Rows[0].Cells[0].Value.ToString()).ToString("000")))
            {
                MessageBoxButtons A;
                A = MessageBoxButtons.OK;
                if (Environment.UserName.ToLower().Contains("rui") || Environment.UserName.ToLower().Contains("pedro"))
                {
                    A = MessageBoxButtons.OKCancel;
                }

                DialogResult a = MessageBox.Show(this, "Impossivel Mover Ficheiro. A pasta já foi criada.", "Error", A, MessageBoxIcon.Error);
                if (a == DialogResult.OK)
                {
                    PASTA = false;
                }
            }


            if (PASTA)
            {
                for (int i = 0; i < dataGridView1.Rows.Count - 1; i++)
                {
                    if (dataGridView1.Rows[i].Cells[5].Value.ToString().Replace(" ", "") == "CQ")
                    {
                        moveCQ(dataGridView1.Rows[i].Cells[0].Value.ToString().Replace(" ", ""), dataGridView1.Rows[i].Cells[1].Value.ToString().Replace(" ", ""), dataGridView1.Rows[i].Cells[2].Value.ToString().Replace(" ", ""), dataGridView1.Rows[i].Cells[3].Value.ToString().Replace(" ", ""), dataGridView1.Rows[i].Cells[4].Value.ToString().Replace(" ", ""));
                    }
                    else if (dataGridView1.Rows[i].Cells[5].Value.ToString().Replace(" ", "") == "CL")
                    {
                        movelaser(dataGridView1.Rows[i].Cells[0].Value.ToString().Replace(" ", ""), dataGridView1.Rows[i].Cells[1].Value.ToString().Replace(" ", ""), dataGridView1.Rows[i].Cells[2].Value.ToString().Replace(" ", ""), dataGridView1.Rows[i].Cells[3].Value.ToString().Replace(" ", ""), dataGridView1.Rows[i].Cells[4].Value.ToString().Replace(" ", ""));

                    }
                    else if (dataGridView1.Rows[i].Cells[5].Value.ToString().Replace(" ", "") == "CM")
                    {
                        moveoxicorte(dataGridView1.Rows[i].Cells[0].Value.ToString().Replace(" ", ""), dataGridView1.Rows[i].Cells[1].Value.ToString().Replace(" ", ""), dataGridView1.Rows[i].Cells[2].Value.ToString().Replace(" ", ""), dataGridView1.Rows[i].Cells[3].Value.ToString().Replace(" ", ""), dataGridView1.Rows[i].Cells[4].Value.ToString().Replace(" ", ""));

                    }
                    else
                    {

                        if (dataGridView1.Rows[i].Cells[6].Value.ToString().ToLower().Replace(" ", "") == "corteefuração")
                        {
                            moveCentralDeCorte(dataGridView1.Rows[i].Cells[0].Value.ToString().Replace(" ", ""), dataGridView1.Rows[i].Cells[1].Value.ToString().Replace(" ", ""), dataGridView1.Rows[i].Cells[2].Value.ToString().Replace(" ", ""), dataGridView1.Rows[i].Cells[3].Value.ToString().Replace(" ", ""), dataGridView1.Rows[i].Cells[4].Value.ToString().Replace(" ", ""));

                        }
                        else if (dataGridView1.Rows[i].Cells[6].Value.ToString().Replace(" ", "") == "Corte")
                        {
                            moveCorte(dataGridView1.Rows[i].Cells[0].Value.ToString().Replace(" ", ""), dataGridView1.Rows[i].Cells[1].Value.ToString().Replace(" ", ""), dataGridView1.Rows[i].Cells[2].Value.ToString().Replace(" ", ""), dataGridView1.Rows[i].Cells[3].Value.ToString().Replace(" ", ""), dataGridView1.Rows[i].Cells[4].Value.ToString().Replace(" ", ""));
                        }

                        if (dataGridView1.Rows[i].Cells[5].Value.ToString().Replace(" ", "") == "CP" && File.Exists(@"c:\R\" + _FrmPai.CBunidadenegocio.Text + "." + _FrmPai.lbl_numeroobra.Text + "." + dataGridView1.Rows[i].Cells[2].Value.ToString().Replace(" ", "") + ".pdf"))
                        {
                            COPYCP(dataGridView1.Rows[i].Cells[0].Value.ToString().Replace(" ", ""), dataGridView1.Rows[i].Cells[1].Value.ToString().Replace(" ", ""), dataGridView1.Rows[i].Cells[2].Value.ToString().Replace(" ", ""), dataGridView1.Rows[i].Cells[3].Value.ToString().Replace(" ", ""), dataGridView1.Rows[i].Cells[4].Value.ToString().Replace(" ", ""));

                        }
                    }
                }

                dataGridView1.Columns.Add("resultado", "resultado");
                for (int i = 0; i < dataGridView1.Rows.Count - 1; i++)
                {
                    if (dataGridView1.Rows[i].Cells[5].Value.ToString().Replace(" ", "") == "CQ")
                    {
                        movechapacm(dataGridView1.Rows[i].Cells[0].Value.ToString().Replace(" ", ""), dataGridView1.Rows[i].Cells[1].Value.ToString().Replace(" ", ""), dataGridView1.Rows[i].Cells[2].Value.ToString().Replace(" ", ""), dataGridView1.Rows[i].Cells[3].Value.ToString().Replace(" ", ""), dataGridView1.Rows[i].Cells[4].Value.ToString().Replace(" ", ""), i);
                        dataGridView1.Rows[i].Cells[7].Value = "feito";


                    }
                    else if (dataGridView1.Rows[i].Cells[5].Value.ToString().Replace(" ", "") == "CL" || dataGridView1.Rows[i].Cells[5].Value.ToString().Replace(" ", "") == "CM")
                    {
                        movechapacm(dataGridView1.Rows[i].Cells[0].Value.ToString().Replace(" ", ""), dataGridView1.Rows[i].Cells[1].Value.ToString().Replace(" ", ""), dataGridView1.Rows[i].Cells[2].Value.ToString().Replace(" ", ""), dataGridView1.Rows[i].Cells[3].Value.ToString().Replace(" ", ""), dataGridView1.Rows[i].Cells[4].Value.ToString().Replace(" ", ""), i);
                        dataGridView1.Rows[i].Cells[7].Value = "feito";
                    }
                    else
                    {
                        if (dataGridView1.Rows[i].Cells[6].Value.ToString().Replace(" ", "") == "CorteeFuração")
                        {

                            dataGridView1.Rows[i].Cells[7].Value = "feito";

                        }
                        else if (dataGridView1.Rows[i].Cells[6].Value.ToString().Replace(" ", "") == "Corte")
                        {

                            dataGridView1.Rows[i].Cells[7].Value = "feito";
                        }
                        else if (dataGridView1.Rows[i].Cells[6].Value.ToString().Replace(" ", "").Contains("Opção"))
                        {
                            movearmacao(dataGridView1.Rows[i].Cells[0].Value.ToString().Replace(" ", ""), dataGridView1.Rows[i].Cells[1].Value.ToString().Replace(" ", ""), dataGridView1.Rows[i].Cells[2].Value.ToString().Replace(" ", ""), dataGridView1.Rows[i].Cells[3].Value.ToString().Replace(" ", ""), dataGridView1.Rows[i].Cells[4].Value.ToString().Replace(" ", ""), dataGridView1.Rows[i].Cells[6].Value.ToString().Replace(" ", ""));
                            dataGridView1.Rows[i].Cells[7].Value = "feito";
                        }
                    }                  
                }
                dumpficheiros(_Formprin.fase);
                _FrmPai.button222_Click(sender, e);
                string fase = int.Parse(dataGridView1.Rows[0].Cells[0].Value.ToString()).ToString("000");
                Excel(fase);
            }           

            AppAbrirPrimavera primaveraHandler = new AppAbrirPrimavera();
            primaveraHandler.AbrirPrimaveira();
        }

      public void dumpficheiros(string fase)
        {
            foreach (string ficheiro in Directory.GetFiles(@"c:\r\"))
            {
                if (Directory.Exists(Frm_Inico.PastaReservatorioFicheiros + int.Parse(dataGridView1.Rows[0].Cells[0].Value.ToString()).ToString("000") + @"\Outros"))
                {
                    File.Move(ficheiro, Frm_Inico.PastaReservatorioFicheiros + int.Parse(dataGridView1.Rows[0].Cells[0].Value.ToString()).ToString("000") + @"\Outros\" + ficheiro.Split('\\').Last());
                }
                else
                {
                    Directory.CreateDirectory(Frm_Inico.PastaReservatorioFicheiros + int.Parse(dataGridView1.Rows[0].Cells[0].Value.ToString()).ToString("000") + @"\Outros");
                    File.Move(ficheiro, Frm_Inico.PastaReservatorioFicheiros + int.Parse(dataGridView1.Rows[0].Cells[0].Value.ToString()).ToString("000") + @"\Outros\" + ficheiro.Split('\\').Last());
                }
            }

            Directory.CreateDirectory(Frm_Inico.PastaReservatorioFicheiros + int.Parse(dataGridView1.Rows[0].Cells[0].Value.ToString()).ToString("000") + "\\2 nesting\\N\\20002");
            Directory.CreateDirectory(Frm_Inico.PastaReservatorioFicheiros + int.Parse(dataGridView1.Rows[0].Cells[0].Value.ToString()).ToString("000") + "\\2 nesting\\N\\20003");
            Directory.CreateDirectory(Frm_Inico.PastaReservatorioFicheiros + int.Parse(dataGridView1.Rows[0].Cells[0].Value.ToString()).ToString("000") + "\\2 nesting\\Q\\20002");
            Directory.CreateDirectory(Frm_Inico.PastaReservatorioFicheiros + int.Parse(dataGridView1.Rows[0].Cells[0].Value.ToString()).ToString("000") + "\\2 nesting\\Q\\20003");

        }

        public void movechapacm(string fase, string lote, string nome, string perfil, string material, int linhaatual)
        {
            string caminhopeca = Frm_Inico.PastaReservatorioFicheiros + int.Parse(dataGridView1.Rows[0].Cells[0].Value.ToString()).ToString("000") + "\\20001\\";

            string VALIDA = "NAO";
            do
            {

                /////////////////////////////////////////////
                if (File.Exists(@"c:\R\" + _FrmPai.CBunidadenegocio.Text + "." + _FrmPai.lbl_numeroobra.Text + "." + nome + ".pdf"))
                {
                    if (!Directory.Exists(caminhopeca))
                    {
                        Directory.CreateDirectory(caminhopeca);
                    }
                    if (!File.Exists(caminhopeca + "\\" + _FrmPai.CBunidadenegocio.Text + "." + _FrmPai.lbl_numeroobra.Text + "." + nome + ".pdf"))
                    {

                        File.Move(@"c:\R\" + _FrmPai.CBunidadenegocio.Text + "." + _FrmPai.lbl_numeroobra.Text + "." + nome + ".pdf", caminhopeca + "\\" + _FrmPai.CBunidadenegocio.Text + "." + _FrmPai.lbl_numeroobra.Text + "." + nome + ".pdf");

                        if (File.Exists(@"c:\R\" + nome + ".dxf"))
                        {
                            File.Delete(@"c:\R\" + nome + ".dxf");
                        }

                    }
                    else
                    {
                        DialogResult RE = MessageBox.Show(this, "Já existe o ficheiro em :" + caminhopeca + "\\" + _FrmPai.CBunidadenegocio.Text + "." + _FrmPai.lbl_numeroobra.Text + "." + nome + ".pdf" + Environment.NewLine + "Deseja substituir", "erro", MessageBoxButtons.YesNo, MessageBoxIcon.Error);
                        if (RE == DialogResult.Yes)
                        {
                            File.Delete(caminhopeca + "\\" + _FrmPai.CBunidadenegocio.Text + "." + _FrmPai.lbl_numeroobra.Text + "." + nome + ".pdf");
                            File.Move(@"c:\R\" + _FrmPai.CBunidadenegocio.Text + "." + _FrmPai.lbl_numeroobra.Text + "." + nome + ".pdf", caminhopeca + "\\" + _FrmPai.CBunidadenegocio.Text + "." + _FrmPai.lbl_numeroobra.Text + "." + nome + ".pdf");
                            if (File.Exists(@"c:\R\" + nome + ".dxf"))
                            {
                                File.Delete(@"c:\R\" + nome + ".dxf");
                            }
                        }
                    }
                    VALIDA = "SIM";
                }
                else
                {
                    if (!File.Exists(caminhopeca + "\\" + _FrmPai.CBunidadenegocio.Text + "." + _FrmPai.lbl_numeroobra.Text + "." + nome + ".pdf"))
                    {

                        //AESPERADOSFICHEIROS F = new AESPERADOSFICHEIROS(@"c:\R\2." + _FrmPai.lbl_numeroobra.Text + "." + nome + ".pdf");
                        //F.ShowDialog();
                        //if (AESPERADOSFICHEIROS.cancel == true)
                        //{
                        VALIDA = "SIM";
                        //}
                    }
                    else
                    {
                        VALIDA = "SIM";
                    }
                }

            } while (VALIDA == "NAO" || dataGridView1.Rows.Count - 1 == linhaatual);
        }

        public void movearmacao(string fase, string lote, string nome, string perfil, string material, string OPCAO)
        {

            string caminhopeca = Frm_Inico.PastaReservatorioFicheiros + int.Parse(dataGridView1.Rows[0].Cells[0].Value.ToString()).ToString("000") + "\\20004";

            string VALIDA = "NAO";
            do
            {

                if (File.Exists(@"c:\R\" + _FrmPai.CBunidadenegocio.Text + "." + _FrmPai.lbl_numeroobra.Text + "." + nome + ".pdf"))
                {
                    if (!Directory.Exists(caminhopeca))
                    {
                        Directory.CreateDirectory(caminhopeca);
                    }
                    if (!File.Exists(caminhopeca + "\\" + _FrmPai.CBunidadenegocio.Text + "." + _FrmPai.lbl_numeroobra.Text + "." + nome + ".pdf"))
                    {
                        File.Move(@"c:\R\" + _FrmPai.CBunidadenegocio.Text + "." + _FrmPai.lbl_numeroobra.Text + "." + nome + ".pdf", caminhopeca + "\\" + _FrmPai.CBunidadenegocio.Text + "." + _FrmPai.lbl_numeroobra.Text + "." + nome + ".pdf");
                    }
                    else
                    {
                        DialogResult RE = MessageBox.Show(this, "Já existe o ficheiro em :" + caminhopeca + "\\" + _FrmPai.CBunidadenegocio.Text + "." + _FrmPai.lbl_numeroobra.Text + "." + nome + ".pdf" + Environment.NewLine + "Deseja substituir", "erro", MessageBoxButtons.YesNo, MessageBoxIcon.Error);
                        if (RE == DialogResult.Yes)
                        {
                            File.Delete(caminhopeca + "\\" + _FrmPai.CBunidadenegocio.Text + "." + _FrmPai.lbl_numeroobra.Text + "." + nome + ".pdf");
                            File.Move(@"c:\R\" + _FrmPai.CBunidadenegocio.Text + "." + _FrmPai.lbl_numeroobra.Text + "." + nome + ".pdf", caminhopeca + "\\" + _FrmPai.CBunidadenegocio.Text + "." + _FrmPai.lbl_numeroobra.Text + "." + nome + ".pdf");
                        }
                    }
                    VALIDA = "SIM";
                }
                else
                {
                    if (!File.Exists(caminhopeca + "\\" + _FrmPai.CBunidadenegocio.Text + "." + _FrmPai.lbl_numeroobra.Text + "." + nome + ".pdf"))
                    {
                        this.Visible = false;
                        AESPERADOSFICHEIROS F = new AESPERADOSFICHEIROS(@"c:\R\" + _FrmPai.CBunidadenegocio.Text + "." + _FrmPai.lbl_numeroobra.Text + "." + nome + ".pdf");
                        F.ShowDialog();
                        this.Visible = true;
                        if (AESPERADOSFICHEIROS.cancel == true)
                        {
                            VALIDA = "SIM";
                        }
                    }
                    else
                    {
                        VALIDA = "SIM";
                    }
                }

            } while (VALIDA == "NAO");

            VALIDA = "NAO";
            caminhopeca = Frm_Inico.PastaReservatorioFicheiros + int.Parse(dataGridView1.Rows[0].Cells[0].Value.ToString()).ToString("000") + "\\20005";

            do
            {

                if (OPCAO.ToLower() == "opção2" || OPCAO.ToLower() == "opção5" || OPCAO.ToLower() == "opção6" || OPCAO.ToLower() == "opção16")
                {
                    if (File.Exists(@"c:\R\" + _FrmPai.CBunidadenegocio.Text + "." + _FrmPai.lbl_numeroobra.Text + "." + nome + " - 1.pdf"))
                    {
                        if (!Directory.Exists(caminhopeca))
                        {
                            Directory.CreateDirectory(caminhopeca);
                        }
                        if (!File.Exists(caminhopeca + "\\" + _FrmPai.CBunidadenegocio.Text + "." + _FrmPai.lbl_numeroobra.Text + "." + nome + " - 1.pdf"))
                        {
                            File.Move(@"c:\R\" + _FrmPai.CBunidadenegocio.Text + "." + _FrmPai.lbl_numeroobra.Text + "." + nome + " - 1.pdf", caminhopeca + "\\" + _FrmPai.CBunidadenegocio.Text + "." + _FrmPai.lbl_numeroobra.Text + "." + nome + " - 1.pdf");
                        }
                        else
                        {
                            DialogResult RE = MessageBox.Show(this, "Já existe o ficheiro em :" + caminhopeca + "\\" + _FrmPai.CBunidadenegocio.Text + "." + _FrmPai.lbl_numeroobra.Text + "." + nome + " - 1.pdf" + Environment.NewLine + "Deseja substituir", "erro", MessageBoxButtons.YesNo, MessageBoxIcon.Error);
                            if (RE == DialogResult.Yes)
                            {
                                File.Delete(caminhopeca + "\\" + _FrmPai.CBunidadenegocio.Text + "." + _FrmPai.lbl_numeroobra.Text + "." + nome + " - 1.pdf");
                                File.Move(@"c:\R\" + _FrmPai.CBunidadenegocio.Text + "." + _FrmPai.lbl_numeroobra.Text + "." + nome + ".pdf", caminhopeca + "\\" + _FrmPai.CBunidadenegocio.Text + "." + _FrmPai.lbl_numeroobra.Text + "." + nome + " - 1.pdf");
                            }
                        }
                        VALIDA = "SIM";
                        string[] files = Directory.GetFiles("c:\\r\\", "*.pdf");


                        foreach (var item in files)
                        {
                            if (item.Contains("Lista de soldadores") || item.Contains("Plano_Soldadura"))
                            {
                                File.Move(item, caminhopeca + "\\" + item.Replace("c:\\r\\", ""));
                            }
                        }
                    }
                    else
                    {
                        MessageBox.Show(this, "FALTA O DESENHO DE SOLDADURA DO CONJUNTO " + nome);
                        VALIDA = "SIM";
                    }
                }
                else
                {
                    VALIDA = "SIM";
                }

            } while (VALIDA == "NAO");

            if (OPCAO.ToLower() == "opção3" || OPCAO.ToLower() == "opção4" || OPCAO.ToLower() == "opção5" || OPCAO.ToLower() == "opção6" || OPCAO.ToLower() == "opção15" || OPCAO.ToLower() == "opção16")
            {
                if (!Directory.Exists(Frm_Inico.PastaReservatorioFicheiros + int.Parse(dataGridView1.Rows[0].Cells[0].Value.ToString()).ToString("000") + "\\20006"))
                {
                    Directory.CreateDirectory(Frm_Inico.PastaReservatorioFicheiros + int.Parse(dataGridView1.Rows[0].Cells[0].Value.ToString()).ToString("000") + "\\20006");
                }
                if (!Directory.Exists(Frm_Inico.PastaReservatorioFicheiros + int.Parse(dataGridView1.Rows[0].Cells[0].Value.ToString()).ToString("000") + "\\20007"))
                {
                    Directory.CreateDirectory(Frm_Inico.PastaReservatorioFicheiros + int.Parse(dataGridView1.Rows[0].Cells[0].Value.ToString()).ToString("000") + "\\20007");
                }
                if (!Directory.Exists(Frm_Inico.PastaReservatorioFicheiros + int.Parse(dataGridView1.Rows[0].Cells[0].Value.ToString()).ToString("000") + "\\20009"))
                {
                    Directory.CreateDirectory(Frm_Inico.PastaReservatorioFicheiros + int.Parse(dataGridView1.Rows[0].Cells[0].Value.ToString()).ToString("000") + "\\20009");
                }
            }
            if (OPCAO.ToLower() == "opção8" || OPCAO.ToLower() == "opção2" || OPCAO.ToLower() == "opção1")
            {
                if (!Directory.Exists(Frm_Inico.PastaReservatorioFicheiros + int.Parse(dataGridView1.Rows[0].Cells[0].Value.ToString()).ToString("000") + "\\20009"))
                {
                    Directory.CreateDirectory(Frm_Inico.PastaReservatorioFicheiros + int.Parse(dataGridView1.Rows[0].Cells[0].Value.ToString()).ToString("000") + "\\20009");
                }
            }
        }

        public void moveCorte(string fase, string lote, string nome, string perfil, string material)
        {
            string caminhopeca = Frm_Inico.PastaReservatorioFicheiros + int.Parse(dataGridView1.Rows[0].Cells[0].Value.ToString()).ToString("000") + "\\20002\\" + perfil;

            string VALIDA = "NAO";
            do
            {
                if (!File.Exists(caminhopeca + "\\" + _FrmPai.CBunidadenegocio.Text + "." + _FrmPai.lbl_numeroobra.Text + "." + nome + ".pdf"))
                {
                    if (File.Exists(@"c:\R\" + _FrmPai.CBunidadenegocio.Text + "." + _FrmPai.lbl_numeroobra.Text + "." + nome + ".pdf"))
                    {
                        if (!Directory.Exists(caminhopeca))
                        {
                            Directory.CreateDirectory(caminhopeca);
                        }
                        if (!File.Exists(caminhopeca + "\\" + _FrmPai.CBunidadenegocio.Text + "." + _FrmPai.lbl_numeroobra.Text + "." + nome + ".pdf"))
                        {
                            File.Move(@"c:\R\" + _FrmPai.CBunidadenegocio.Text + "." + _FrmPai.lbl_numeroobra.Text + "." + nome + ".pdf", caminhopeca + "\\" + _FrmPai.CBunidadenegocio.Text + "." + _FrmPai.lbl_numeroobra.Text + "." + nome + ".pdf");
                        }
                        else
                        {
                            DialogResult RE = MessageBox.Show(this, "Já existe o ficheiro em :" + caminhopeca + "\\" + _FrmPai.CBunidadenegocio.Text + "." + _FrmPai.lbl_numeroobra.Text + "." + nome + ".pdf" + Environment.NewLine + "Deseja substituir", "erro", MessageBoxButtons.YesNo, MessageBoxIcon.Error);
                            if (RE == DialogResult.Yes)
                            {
                                File.Delete(caminhopeca + "\\" + _FrmPai.CBunidadenegocio.Text + "." + _FrmPai.lbl_numeroobra.Text + "." + nome + ".pdf");
                                File.Move(@"c:\R\" + _FrmPai.CBunidadenegocio.Text + "." + _FrmPai.lbl_numeroobra.Text + "." + nome + ".pdf", caminhopeca + "\\" + _FrmPai.CBunidadenegocio.Text + "." + _FrmPai.lbl_numeroobra.Text + "." + nome + ".pdf");
                            }
                        }
                        VALIDA = "SIM";
                    }
                    else
                    {
                        if (!File.Exists(caminhopeca + "\\" + _FrmPai.CBunidadenegocio.Text + "." + _FrmPai.lbl_numeroobra.Text + "." + nome + ".pdf"))
                        {
                            this.Visible = false;
                            AESPERADOSFICHEIROS F = new AESPERADOSFICHEIROS(@"c:\R\" + _FrmPai.CBunidadenegocio.Text + "." + _FrmPai.lbl_numeroobra.Text + "." + nome + ".pdf");
                            F.ShowDialog();
                            this.Visible = true;
                            if (AESPERADOSFICHEIROS.cancel == true)
                            {
                                VALIDA = "SIM";
                            }
                        }
                        else
                        {
                            VALIDA = "SIM";
                        }
                    }
                }
                else
                {
                    VALIDA = "SIM";
                }
            } while (VALIDA == "NAO");
        }

        public void moveCentralDeCorte(string fase, string lote, string nome, string perfil, string material)
        {
            string caminhonc = @"\\vernet\prod\Obras " + Frm_Inico.ano + "\\" + _FrmPai.lbl_numeroobra.Text + "\\" + int.Parse(fase).ToString("000");
            string caminhopeca = Frm_Inico.PastaReservatorioFicheiros + int.Parse(dataGridView1.Rows[0].Cells[0].Value.ToString()).ToString("000") + "\\20003\\" + perfil;

            string VALIDA = "NAO";
            do
            {

                if (File.Exists(@"c:\R\" + _FrmPai.CBunidadenegocio.Text + "." + _FrmPai.lbl_numeroobra.Text + "." + nome + ".pdf"))
                {
                    if (!Directory.Exists(caminhopeca))
                    {
                        Directory.CreateDirectory(caminhopeca);
                    }
                    if (!File.Exists(caminhopeca + "\\" + _FrmPai.CBunidadenegocio.Text + "." + _FrmPai.lbl_numeroobra.Text + "." + nome + ".pdf"))
                    {
                        File.Move(@"c:\R\" + _FrmPai.CBunidadenegocio.Text + "." + _FrmPai.lbl_numeroobra.Text + "." + nome + ".pdf", caminhopeca + "\\" + _FrmPai.CBunidadenegocio.Text + "." + _FrmPai.lbl_numeroobra.Text + "." + nome + ".pdf");
                    }
                    else
                    {
                        DialogResult RE = MessageBox.Show(this, "Já existe o ficheiro em :" + caminhopeca + "\\" + _FrmPai.CBunidadenegocio.Text + "." + _FrmPai.lbl_numeroobra.Text + "." + nome + ".pdf" + Environment.NewLine + "Deseja substituir", "erro", MessageBoxButtons.YesNo, MessageBoxIcon.Error);
                        if (RE == DialogResult.Yes)
                        {
                            File.Delete(caminhopeca + "\\" + _FrmPai.CBunidadenegocio.Text + "." + _FrmPai.lbl_numeroobra.Text + "." + nome + ".pdf");
                            File.Move(@"c:\R\" + _FrmPai.CBunidadenegocio.Text + "." + _FrmPai.lbl_numeroobra.Text + "." + nome + ".pdf", caminhopeca + "\\" + _FrmPai.CBunidadenegocio.Text + "." + _FrmPai.lbl_numeroobra.Text + "." + nome + ".pdf");

                        }
                    }
                    VALIDA = "SIM";
                }
                else
                {
                    if (!File.Exists(caminhopeca + "\\" + _FrmPai.CBunidadenegocio.Text + "." + _FrmPai.lbl_numeroobra.Text + "." + nome + ".pdf"))
                    {
                        this.Visible = false;
                        AESPERADOSFICHEIROS F = new AESPERADOSFICHEIROS(@"c:\R\" + _FrmPai.CBunidadenegocio.Text + "." + _FrmPai.lbl_numeroobra.Text + "." + nome + ".pdf");
                        F.ShowDialog();
                        this.Visible = true;
                        if (AESPERADOSFICHEIROS.cancel == true)
                        {
                            VALIDA = "SIM";
                        }
                    }
                    else
                    {
                        VALIDA = "SIM";
                    }
                }

            } while (VALIDA == "NAO");
        }

        public void moveCQ(string fase, string lote, string nome, string perfil, string material)
        {
            string caminhopeca = Frm_Inico.PastaPartilhada + "\\" + Frm_Inico.ano + @"\CQ\" + _FrmPai.lbl_numeroobra.Text + "\\" + int.Parse(fase).ToString("000");
            string VALIDA = "NAO";
            do
            {
                if (File.Exists(@"c:\R\" + _FrmPai.CBunidadenegocio.Text + "." + _FrmPai.lbl_numeroobra.Text + "." + nome + ".pdf"))
                {


                    if (!Directory.Exists(caminhopeca))
                    {
                        Directory.CreateDirectory(caminhopeca);
                    }
                    if (!File.Exists(caminhopeca + "\\" + _FrmPai.CBunidadenegocio.Text + "." + _FrmPai.lbl_numeroobra.Text + "." + nome + ".pdf"))
                    {
                        File.Copy(@"c:\R\" + _FrmPai.CBunidadenegocio.Text + "." + _FrmPai.lbl_numeroobra.Text + "." + nome + ".pdf", caminhopeca + "\\" + _FrmPai.CBunidadenegocio.Text + "." + _FrmPai.lbl_numeroobra.Text + "." + nome + ".pdf");
                        VALIDA = "SIM";
                    }
                    else
                    {
                        DialogResult RE = MessageBox.Show(this, "Já existe o ficheiro em :" + caminhopeca + "\\" + _FrmPai.CBunidadenegocio.Text + "." + _FrmPai.lbl_numeroobra.Text + "." + nome + ".pdf" + Environment.NewLine + "Deseja substituir", "erro", MessageBoxButtons.YesNo, MessageBoxIcon.Error);
                        if (RE == DialogResult.Yes)
                        {
                            File.Delete(caminhopeca + "\\" + _FrmPai.CBunidadenegocio.Text + "." + _FrmPai.lbl_numeroobra.Text + "." + nome + ".pdf");
                            File.Copy(@"c:\R\" + _FrmPai.CBunidadenegocio.Text + "." + _FrmPai.lbl_numeroobra.Text + "." + nome + ".pdf", caminhopeca + "\\" + _FrmPai.CBunidadenegocio.Text + "." + _FrmPai.lbl_numeroobra.Text + "." + nome + ".pdf");
                            VALIDA = "SIM";
                        }
                    }
                    VALIDA = "SIM";
                }
                else
                {
                    if (!File.Exists(caminhopeca + "\\" + _FrmPai.CBunidadenegocio.Text + "." + _FrmPai.lbl_numeroobra.Text + "." + nome + ".pdf"))
                    {
                        this.Visible = false;
                        AESPERADOSFICHEIROS F = new AESPERADOSFICHEIROS(@"c:\R\" + _FrmPai.CBunidadenegocio.Text + "." + _FrmPai.lbl_numeroobra.Text + "." + nome + ".pdf");
                        F.ShowDialog();
                        this.Visible = true;
                        if (AESPERADOSFICHEIROS.cancel == true)
                        {
                            VALIDA = "SIM";
                        }
                    }
                    else
                    {
                        VALIDA = "SIM";
                    }
                }

            } while (VALIDA == "NAO");
        }

        public void movelaser(string fase, string lote, string nome, string perfil, string material)
        {
            string caminhopeca = Frm_Inico.PastaPartilhada + "\\" + Frm_Inico.ano + @"\LASER\" + _FrmPai.lbl_numeroobra.Text + "\\" + int.Parse(fase).ToString("000") + "\\" + perfil + "_" + material;
            string VALIDA = "NAO";
            do
            {
                if (File.Exists(@"c:\R\" + _FrmPai.CBunidadenegocio.Text + "." + _FrmPai.lbl_numeroobra.Text + "." + nome + ".pdf"))
                {
                    if (!Directory.Exists(caminhopeca))
                    {
                        Directory.CreateDirectory(caminhopeca);
                    }
                    if (!File.Exists(caminhopeca + "\\" + _FrmPai.CBunidadenegocio.Text + "." + _FrmPai.lbl_numeroobra.Text + "." + nome + ".pdf"))
                    {
                        File.Copy(@"c:\R\" + _FrmPai.CBunidadenegocio.Text + "." + _FrmPai.lbl_numeroobra.Text + "." + nome + ".pdf", caminhopeca + "\\" + _FrmPai.CBunidadenegocio.Text + "." + _FrmPai.lbl_numeroobra.Text + "." + nome + ".pdf");

                    }
                    else
                    {
                        DialogResult RE = MessageBox.Show(this, "Já existe o ficheiro em :" + caminhopeca + "\\" + _FrmPai.CBunidadenegocio.Text + "." + _FrmPai.lbl_numeroobra.Text + "." + nome + ".pdf" + Environment.NewLine + "Deseja substituir", "erro", MessageBoxButtons.YesNo, MessageBoxIcon.Error);
                        if (RE == DialogResult.Yes)
                        {
                            File.Delete(caminhopeca + "\\" + _FrmPai.CBunidadenegocio.Text + "." + _FrmPai.lbl_numeroobra.Text + "." + nome + ".pdf");
                        }
                    }

                    VALIDA = "SIM";
                }
                else
                {
                    if (!File.Exists(caminhopeca + "\\" + _FrmPai.CBunidadenegocio.Text + "." + _FrmPai.lbl_numeroobra.Text + "." + nome + ".pdf"))
                    {
                        this.Visible = false;
                        AESPERADOSFICHEIROS F = new AESPERADOSFICHEIROS(@"c:\R\" + _FrmPai.CBunidadenegocio.Text + "." + _FrmPai.lbl_numeroobra.Text + "." + nome + ".pdf");
                        F.ShowDialog();
                        this.Visible = true;
                        if (AESPERADOSFICHEIROS.cancel == true)
                        {
                            VALIDA = "SIM";
                        }
                    }
                    else
                    {
                        VALIDA = "SIM";
                    }
                }
            } while (VALIDA == "NAO");

            VALIDA = "NAO";
            do
            {
                if (File.Exists(@"c:\R\" + nome + ".dxf"))
                {

                    if (!File.Exists(caminhopeca + "\\" + _FrmPai.CBunidadenegocio.Text + "." + _FrmPai.lbl_numeroobra.Text + "." + nome + ".dxf"))
                    {
                        File.Copy(@"c:\R\" + nome + ".dxf", caminhopeca + "\\" + _FrmPai.CBunidadenegocio.Text + "." + _FrmPai.lbl_numeroobra.Text + "." + nome + ".dxf");

                    }
                    else
                    {
                        DialogResult RE = MessageBox.Show(this, "Já existe o ficheiro em :" + caminhopeca + "\\" + _FrmPai.CBunidadenegocio.Text + "." + _FrmPai.lbl_numeroobra.Text + "." + nome + ".dxf" + Environment.NewLine + "Deseja substituir", "erro", MessageBoxButtons.YesNo, MessageBoxIcon.Error);
                        if (RE == DialogResult.Yes)
                        {
                            File.Delete(caminhopeca + "\\" + _FrmPai.CBunidadenegocio.Text + "." + _FrmPai.lbl_numeroobra.Text + "." + nome + ".dxf");
                            File.Copy(@"c:\R\" + nome + ".dxf", caminhopeca + "\\" + _FrmPai.CBunidadenegocio.Text + "." + _FrmPai.lbl_numeroobra.Text + "." + nome + ".dxf");

                        }
                    }

                    VALIDA = "SIM";
                }
                else
                {
                    if (!File.Exists(caminhopeca + "\\" + _FrmPai.CBunidadenegocio.Text + "." + _FrmPai.lbl_numeroobra.Text + "." + nome + ".dxf"))
                    {
                        this.Visible = false;
                        AESPERADOSFICHEIROS F = new AESPERADOSFICHEIROS(@"c:\R\" + _FrmPai.CBunidadenegocio.Text + "." + _FrmPai.lbl_numeroobra.Text + "." + nome + ".dxf");
                        F.ShowDialog();
                        this.Visible = true;
                        if (AESPERADOSFICHEIROS.cancel == true)
                        {
                            VALIDA = "SIM";
                        }
                    }
                    else
                    {
                        VALIDA = "SIM";
                    }
                }
            } while (VALIDA == "NAO");
        }

        public void moveoxicorte(string fase, string lote, string nome, string perfil, string material)
        {
            string caminhopeca = Frm_Inico.PastaPartilhada + "\\" + Frm_Inico.ano + @"\CM\" + _FrmPai.lbl_numeroobra.Text + "\\" + int.Parse(fase).ToString("000") + "\\" + perfil + "_" + material;
            string VALIDA = "NAO";
            do
            {
                if (File.Exists(@"c:\R\" + _FrmPai.CBunidadenegocio.Text + "." + _FrmPai.lbl_numeroobra.Text + "." + nome + ".pdf"))
                {
                    if (!Directory.Exists(caminhopeca))
                    {
                        Directory.CreateDirectory(caminhopeca);
                    }
                    if (!File.Exists(caminhopeca + "\\" + _FrmPai.CBunidadenegocio.Text + "." + _FrmPai.lbl_numeroobra.Text + "." + nome + ".pdf"))
                    {
                        File.Copy(@"c:\R\" + _FrmPai.CBunidadenegocio.Text + "." + _FrmPai.lbl_numeroobra.Text + "." + nome + ".pdf", caminhopeca + "\\" + _FrmPai.CBunidadenegocio.Text + "." + _FrmPai.lbl_numeroobra.Text + "." + nome + ".pdf");

                    }
                    else
                    {
                        DialogResult RE = MessageBox.Show(this, "Já existe o ficheiro em :" + caminhopeca + "\\" + _FrmPai.CBunidadenegocio.Text + "." + _FrmPai.lbl_numeroobra.Text + "." + nome + ".pdf" + Environment.NewLine + "Deseja substituir", "erro", MessageBoxButtons.YesNo, MessageBoxIcon.Error);
                        if (RE == DialogResult.Yes)
                        {
                            File.Delete(caminhopeca + "\\" + _FrmPai.CBunidadenegocio.Text + "." + _FrmPai.lbl_numeroobra.Text + "." + nome + ".pdf");
                        }
                    }

                    VALIDA = "SIM";
                }
                else
                {
                    if (!File.Exists(caminhopeca + "\\" + _FrmPai.CBunidadenegocio.Text + "." + _FrmPai.lbl_numeroobra.Text + "." + nome + ".pdf"))
                    {
                        this.Visible = false;
                        AESPERADOSFICHEIROS F = new AESPERADOSFICHEIROS(@"c:\R\" + _FrmPai.CBunidadenegocio.Text + "." + _FrmPai.lbl_numeroobra.Text + "." + nome + ".pdf");
                        F.ShowDialog();
                        this.Visible = true;
                        if (AESPERADOSFICHEIROS.cancel == true)
                        {
                            VALIDA = "SIM";
                        }
                    }
                    else
                    {
                        VALIDA = "SIM";
                    }
                }
            } while (VALIDA == "NAO");

            VALIDA = "NAO";
            do
            {
                if (File.Exists(@"c:\R\" + nome + ".dxf"))
                {

                    if (!File.Exists(caminhopeca + "\\" + _FrmPai.CBunidadenegocio.Text + "." + _FrmPai.lbl_numeroobra.Text + "." + nome + ".dxf"))
                    {
                        File.Copy(@"c:\R\" + nome + ".dxf", caminhopeca + "\\" + _FrmPai.CBunidadenegocio.Text + "." + _FrmPai.lbl_numeroobra.Text + "." + nome + ".dxf");

                    }
                    else
                    {
                        DialogResult RE = MessageBox.Show(this, "Já existe o ficheiro em :" + caminhopeca + "\\" + _FrmPai.CBunidadenegocio.Text + "." + _FrmPai.lbl_numeroobra.Text + "." + nome + ".dxf" + Environment.NewLine + "Deseja substituir", "erro", MessageBoxButtons.YesNo, MessageBoxIcon.Error);
                        if (RE == DialogResult.Yes)
                        {
                            File.Delete(caminhopeca + "\\" + _FrmPai.CBunidadenegocio.Text + "." + _FrmPai.lbl_numeroobra.Text + "." + nome + ".dxf");
                            File.Copy(@"c:\R\" + nome + ".dxf", caminhopeca + "\\" + _FrmPai.CBunidadenegocio.Text + "." + _FrmPai.lbl_numeroobra.Text + "." + nome + ".dxf");

                        }
                    }

                    VALIDA = "SIM";
                }
                else
                {
                    if (!File.Exists(caminhopeca + "\\" + _FrmPai.CBunidadenegocio.Text + "." + _FrmPai.lbl_numeroobra.Text + "." + nome + ".dxf"))
                    {
                        this.Visible = false;
                        AESPERADOSFICHEIROS F = new AESPERADOSFICHEIROS(@"c:\R\" + _FrmPai.CBunidadenegocio.Text + "." + _FrmPai.lbl_numeroobra.Text + "." + nome + ".dxf");
                        F.ShowDialog();
                        this.Visible = true;
                        if (AESPERADOSFICHEIROS.cancel == true)
                        {
                            VALIDA = "SIM";
                        }
                    }
                    else
                    {
                        VALIDA = "SIM";
                    }
                }
            } while (VALIDA == "NAO");
        }

        public void COPYCP(string fase, string lote, string nome, string perfil, string material)
        {
            string caminhopeca = Frm_Inico.PastaPartilhada + "\\" + Frm_Inico.ano + @"\CP\" + _FrmPai.lbl_numeroobra.Text + "\\" + int.Parse(fase).ToString("000") + "\\" + perfil;


            if (File.Exists(@"c:\R\" + _FrmPai.CBunidadenegocio.Text + "." + _FrmPai.lbl_numeroobra.Text + "." + nome + ".pdf"))
            {
                if (!Directory.Exists(caminhopeca))
                {
                    Directory.CreateDirectory(caminhopeca);
                }
                if (!File.Exists(caminhopeca + "\\" + _FrmPai.CBunidadenegocio.Text + "." + _FrmPai.lbl_numeroobra.Text + "." + nome + ".pdf"))
                {
                    File.Copy(@"c:\R\" + _FrmPai.CBunidadenegocio.Text + "." + _FrmPai.lbl_numeroobra.Text + "." + nome + ".pdf", caminhopeca + "\\" + _FrmPai.CBunidadenegocio.Text + "." + _FrmPai.lbl_numeroobra.Text + "." + nome + ".pdf");

                }
                else
                {
                    DialogResult RE = MessageBox.Show(this, "Já existe o ficheiro em :" + caminhopeca + "\\" + _FrmPai.CBunidadenegocio.Text + "." + _FrmPai.lbl_numeroobra.Text + "." + nome + ".pdf" + Environment.NewLine + "Deseja substituir", "erro", MessageBoxButtons.YesNo, MessageBoxIcon.Error);
                    if (RE == DialogResult.Yes)
                    {
                        File.Delete(caminhopeca + "\\" + _FrmPai.CBunidadenegocio.Text + "." + _FrmPai.lbl_numeroobra.Text + "." + nome + ".pdf");
                    }
                }
            }

            if (File.Exists(@"c:\R\" + nome + ".nc1"))
            {
                
                if (perfil.Contains("Z") || perfil.Contains("C"))
                {                    
                    if (!Directory.Exists(caminhopeca))
                    {
                        Directory.CreateDirectory(caminhopeca);
                    }
                   
                    if (!File.Exists(caminhopeca + "\\" + _FrmPai.CBunidadenegocio.Text + "." + _FrmPai.lbl_numeroobra.Text + "." + nome + ".nc1"))
                    {
                        File.Copy(@"c:\R\" + nome + ".nc1", caminhopeca + "\\" + _FrmPai.CBunidadenegocio.Text + "." + _FrmPai.lbl_numeroobra.Text + "." + nome + ".nc1");
                    }
                    else
                    {
                        DialogResult RE = MessageBox.Show(this, "Já existe o ficheiro em: " + caminhopeca + "\\" + _FrmPai.CBunidadenegocio.Text + "." + _FrmPai.lbl_numeroobra.Text + "." + nome + ".nc1" + Environment.NewLine + "Deseja substituir?", "Erro", MessageBoxButtons.YesNo, MessageBoxIcon.Error);
                        if (RE == DialogResult.Yes)
                        {
                            File.Delete(caminhopeca + "\\" + _FrmPai.CBunidadenegocio.Text + "." + _FrmPai.lbl_numeroobra.Text + "." + nome + ".nc1");
                            File.Copy(@"c:\R\" + nome + ".nc1", caminhopeca + "\\" + _FrmPai.CBunidadenegocio.Text + "." + _FrmPai.lbl_numeroobra.Text + "." + nome + ".nc1");
                        }
                    }
                }                
            }         

            if (File.Exists(@"c:\R\" + nome + ".nc1"))
            {
                if (!Directory.Exists(caminhopeca))
                {
                    Directory.CreateDirectory(caminhopeca);
                }
                if (!File.Exists(caminhopeca + "\\" + _FrmPai.CBunidadenegocio.Text + "." + _FrmPai.lbl_numeroobra.Text + "." + nome + ".nc1"))
                {
                    File.Copy(@"c:\R\" + nome + ".nc1", caminhopeca + "\\" + _FrmPai.CBunidadenegocio.Text + "." + _FrmPai.lbl_numeroobra.Text + "." + nome + ".nc1");

                }
                else
                {
                    DialogResult RE = MessageBox.Show(this, "Já existe o ficheiro em :" + caminhopeca + "\\" + _FrmPai.CBunidadenegocio.Text + "." + _FrmPai.lbl_numeroobra.Text + "." + nome + ".nc1" + Environment.NewLine + "Deseja substituir", "erro", MessageBoxButtons.YesNo, MessageBoxIcon.Error);
                    if (RE == DialogResult.Yes)
                    {
                        File.Delete(caminhopeca + "\\" + _FrmPai.CBunidadenegocio.Text + "." + _FrmPai.lbl_numeroobra.Text + "." + nome + ".nc1");
                    }
                }
            }
        }
        private void Excel(string fase)
        {
            string caminhocsv = Path.Combine(@"\\marconi\OFELIZ\OFELIZ\OFM\2.AN\2.CM\DP\1 Obras", Frm_Inico.ano, _FrmPai.lbl_numeroobra.Text, "1.9 Gestão de fabrico", int.Parse(fase).ToString("000") + @"\");
            string caminhocsv2 = Path.Combine(Frm_Inico.PastaPartilhada, Frm_Inico.ano, "CP", _FrmPai.lbl_numeroobra.Text, int.Parse(fase).ToString("000"));
            string nomeprojeto = _FrmPai.lbl_numeroobra.Text + "F" + int.Parse(fase).ToString("000");

            string destino = Path.Combine(caminhocsv2, nomeprojeto + ".csv");
            string destinoDir = Path.GetDirectoryName(destino);

          
            try
            {
                File.Copy(caminhocsv + nomeprojeto + ".csv", destino);
            }
            catch (Exception ex)
            {
            }
        }
        private void ChamarPowerfab()
        {
            if (dataGridView1.Rows.Count > 0 && dataGridView1.Rows[0].Cells[0].Value != null)
            {
                if (double.TryParse(dataGridView1.Rows[0].Cells[0].Value.ToString(), out double valor))
                {
                    if (valor < 500)
                    {
                        string numeroObra = _FrmPai.lbl_numeroobra.Text.Trim();
                        CreatePowerfabFabrico(numeroObra);
                        CreatePowerfab(numeroObra);
                        AppAbrirTeklanovo teklaHandler = new AppAbrirTeklanovo();
                        teklaHandler.TrazerTeklaParaFrente();
                    }
                }
            }
        }
        public void CreatePowerfab(string numeroObra)
        {
            string ano = "20" + numeroObra.Substring(0, 2);
            string caminho1 = @"\\marconi\COMPANY SHARED FOLDER\OFELIZ\OFM\2.AN\2.CM\DP\1 Obras\";
            string caminho2 = Path.Combine(caminho1, ano, numeroObra, "1.8 Projeto", "1.8.2 Tekla");

            if (!Directory.Exists(caminho2))
            {
                MessageBox.Show($"A pasta '{caminho2}' não existe.");
                return;
            }

            string[] subpastas = Directory.GetDirectories(caminho2);
            if (subpastas.Length == 0)
            {
                MessageBox.Show($"Nenhuma subpasta encontrada em '{caminho2}'.");
                return;
            }

            string primeiraPasta = subpastas[0];
            string caminho3 = Path.Combine(primeiraPasta, "attributes");
            Directory.CreateDirectory(caminho3);
            string filePath = Path.Combine(caminho3, $"Fabrico.powerfabexport.json");

            if (File.Exists(filePath))
            {
                return;
            }

            var jsonPrincipal = $@" {{
                                      ""DialogType"": ""export.common"",
                                      ""IsEstimatingChecked"": false,
                                      ""IsProcurementChecked"": false,
                                      ""IsDrawingsForApprovalChecked"": false,
                                      ""IsProductionControlChecked"": true,
                                      ""SelectedEstimatingItemName"": ""standard"",
                                      ""SelectedProcurementItemName"": ""standard"",
                                      ""SelectedDrawingsForApprovalItemName"": ""standard"",
                                      ""SelectedProductionControlItemName"": ""Fabrico_Settings"",
                                      ""SelectionTypeIndexForEstimating"": 1,
                                      ""SelectionTypeIndexForProcurement"": 0,
                                      ""SelectionTypeIndexForDrawingsForApproval"": 0,
                                      ""SelectionTypeIndexForProductionControl"": 2,
                                      ""SelectedSharingUploadOptionIndex"": 0,
                                      ""SelectedTrimbleConnectFolder"": """",                                      
                                      ""Folder"": "".\\05Exportado\\Tekla_PowerFab"",
                                      ""TargetFileName"": ""L1_F1_R0"",
                                      ""FileNameSuffix"": """",
                                      ""FileNameTemplateData"": """",
                                      ""LocationByGuid"": ""0ff713ca-2fe9-497a-9534-92e0f76108a2""
                                    }}";
            try
            {
                File.WriteAllText(filePath, jsonPrincipal);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Erro ao criar o ficheiro JSON: {ex.Message}");
            }
        }
        public void CreatePowerfabFabrico(string numeroObra)
        {
            string ano = "20" + numeroObra.Substring(0, 2);
            string caminho1 = @"\\marconi\COMPANY SHARED FOLDER\OFELIZ\OFM\2.AN\2.CM\DP\1 Obras\";
            string caminho2 = Path.Combine(caminho1, ano, numeroObra, "1.8 Projeto", "1.8.2 Tekla");

            if (!Directory.Exists(caminho2))
            {
                MessageBox.Show($"A pasta '{caminho2}' não existe.");
                return;
            }

            string[] subpastas = Directory.GetDirectories(caminho2);
            if (subpastas.Length == 0)
            {
                MessageBox.Show($"Nenhuma subpasta encontrada em '{caminho2}'.");
                return;
            }

            string primeiraPasta = subpastas[0];
            string caminho3 = Path.Combine(primeiraPasta, "attributes");
            Directory.CreateDirectory(caminho3);
            string filePathFabrico = Path.Combine(caminho3, $"Fabrico_Settings.productioncontrol.json");

            if (File.Exists(filePathFabrico))
            {
                return;
            }

            string CaminhoFabrico = $@"\\marconi\\OFELIZ\\OFELIZ\\OFM\\2.AN\\2.CM\\DP\\1 Obras\\{ano}\\{numeroObra}\\1.9 Gestão de fabrico";

            var jsonFabrico = $@"{{
                                    ""DialogType"": ""export.common"",
                                    ""SelectedAssemblyDrawingsIncludeOption"": ""IncludeInfoAndPdf"",
                                    ""SelectedGeneralArrangementDrawingsIncludeOption"": ""DontInclude"",
                                    ""SelectedMultiDrawingsIncludeOption"": ""DontInclude"",
                                    ""SelectedSinglePartDrawingsIncludeOption"": ""IncludeInfoAndPdf"",
                                    ""DrawingUdasIncludeOption"": ""IncludeFromModel"",
                                    ""AssemblyDrawingsFolder"": ""{CaminhoFabrico}"",
                                    ""GeneralArrangementDrawingsFolder"": "".\\01Desenhos\\PDF\\"",
                                    ""MultidrawingsFolder"": "".\\01Desenhos\\PDF\\"",
                                    ""SinglePartDrawingsFolder"": ""{CaminhoFabrico}"",
                                    ""IsDrawingsUDAsChecked"": true,
                                    ""IsPartUDAsChecked"": true,
                                    ""PartUdasIncludeOption"": ""IncludeFromModel"",
                                    ""AssemblyUdasIncludeOption"": ""IncludeFromReport"",
                                    ""BoltNutWasherUdasIncludeOption"": ""IncludeFromModel"",
                                    ""StudUdasIncludeOption"": ""DontInclude"",
                                    ""IsBoltsNutsWashersChecked"": true,
                                    ""IsBoltNutWasherUDAsChecked"": true,
                                    ""IsStudsChecked"": false,
                                    ""IsStudUDAsChecked"": false,
                                    ""IsAssemblyUDAsChecked"": false,
                                    ""IsDoNotExportCNCFilesChecked"": false,
                                    ""IsGenerateCNCFilesSettingsChecked"": false,
                                    ""IsUseCNCFilesFromFolderChecked"": true,
                                    ""UseCNCFilesFolder"": ""{CaminhoFabrico}"",
                                    ""CNCFilesSettingsListItemsIndex"": 0
                                }}";

            try
            {
                File.WriteAllText(filePathFabrico, jsonFabrico);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Erro ao criar o ficheiro JSON: {ex.Message}");
            }
        }
        public class AppAbrirTeklanovo
        {
            [DllImport("user32.dll", SetLastError = true)]
            public static extern bool EnumWindows(EnumWindowsProc enumProc, IntPtr lParam);

            [DllImport("user32.dll", SetLastError = true)]
            public static extern IntPtr GetWindowText(IntPtr hWnd, StringBuilder text, int count);

            [DllImport("user32.dll", SetLastError = true)]
            public static extern IntPtr GetForegroundWindow();

            [DllImport("user32.dll")]
            public static extern bool SetForegroundWindow(IntPtr hWnd);

            [DllImport("user32.dll")]
            public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

            const int SW_RESTORE = 5;

            public delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

            public void TrazerTeklaParaFrente()
            {
                EnumWindows(new EnumWindowsProc(EnumWindowCallback), IntPtr.Zero);
            }

            private bool EnumWindowCallback(IntPtr hWnd, IntPtr lParam)
            {
                StringBuilder windowTitle = new StringBuilder(256);
                GetWindowText(hWnd, windowTitle, 256);

                if (windowTitle.ToString().StartsWith("Tekla Structures"))
                {
                    ShowWindow(hWnd, SW_RESTORE);
                    SetForegroundWindow(hWnd);
                    SimularTeclas();
                    return false;
                }

                return true;
            }

            private void SimularTeclas()
            {
                var simulator = new InputSimulator();
                simulator.Keyboard.ModifiedKeyStroke(new[] { VirtualKeyCode.CONTROL, VirtualKeyCode.SHIFT }, VirtualKeyCode.F11);
            }
        }
        public class AppAbrirPrimavera
        {
            [DllImport("user32.dll")]
            public static extern bool SetForegroundWindow(IntPtr hWnd);

            [DllImport("user32.dll")]
            public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

            const int SW_RESTORE = 9;

            public void AbrirPrimaveira()
            {
                try
                {
                    string nomeProcesso = "Erp100EV";
                    string appPath = @"C:\Program Files (x86)\PRIMAVERA_EVO\SG100\Apl\Erp100EV.exe";

                    Process[] processos = Process.GetProcessesByName(nomeProcesso);

                    if (processos.Length > 0)
                    {
                        Process processoExistente = processos[0];

                        IntPtr hWnd = processoExistente.MainWindowHandle;

                        if (hWnd != IntPtr.Zero)
                        {
                            ShowWindow(hWnd, SW_RESTORE);
                            SetForegroundWindow(hWnd);
                        }
                        else
                        {
                            MessageBox.Show("Primavera já está aberto, mas não foi possível aceder à janela principal.");
                        }
                    }
                    else
                    {
                        if (System.IO.File.Exists(appPath))
                        {
                            Process processoNovo = Process.Start(appPath);

                            Thread.Sleep(1000);

                            IntPtr hWnd = processoNovo.MainWindowHandle;

                            if (hWnd != IntPtr.Zero)
                            {
                                ShowWindow(hWnd, SW_RESTORE);
                                SetForegroundWindow(hWnd);
                            }
                            else
                            {
                                //MessageBox.Show("Primavera iniciado, mas não foi possível aceder à janela principal.");
                            }
                        }
                        else
                        {
                            MessageBox.Show("O Primavera não foi encontrado no PC.");
                        }
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Erro ao tentar abrir o Primavera: " + ex.Message);
                }
            }
        }      
        private void Frm_Pecas_FormClosed(object sender, FormClosedEventArgs e)
        {            
            ChamarPowerfab();
        }      
      }
}