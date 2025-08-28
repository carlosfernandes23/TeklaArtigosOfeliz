using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Runtime.ConstrainedExecution;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace TeklaArtigosOfeliz
{
    public partial class Frm_EditarProRevestimentos : Form
    {
        public string connection;

        string connectionString = @"Data Source=GALILEU\PREPARACAO;Initial Catalog=ArtigoTekla;Persist Security Info=True;User ID=SA;Password=preparacao";

        public Frm_EditarProRevestimentos()
        {
            InitializeComponent();
        }

        private void Frm_EditarProRevestimentos_Load(object sender, EventArgs e)
        {
            ConectarBDcor();
            ConectarBDMaterial();
        }
                

        private void guna2Button2_Click(object sender, EventArgs e)
        {
            InserirCor();
        }       

        private void guna2Button3_Click(object sender, EventArgs e)
        {
            RemoverCorSelecionada();
        }
              

        private void guna2Button4_Click(object sender, EventArgs e)
        {
            InserirMaterial();
        }

        private void guna2Button1_Click(object sender, EventArgs e)
        {
            RemoverMaterialSelecionada();
        }

        private void ConectarBDcor()
        {
            listBoxCor.ClearSelected();
            listBoxCor.Items.Clear();
            ComunicaBDtekla BD = new ComunicaBDtekla();
            string Familia = "Policarbonato";
            try
            {
                BD.ConectarBD();
                List<string> cores = BD.Procurarbd("SELECT [Cor] FROM [dbo].[corpolicarbonato] WHERE [Familia] = '" + Familia.Trim() + "'");

                listBoxCor.Items.Clear();
                foreach (string cor in cores)
                {
                    listBoxCor.Items.Add(cor);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Erro ao buscar dados: " + ex.Message);
            }
            finally
            {
                BD.DesonectarBD();
            }
        }

        private void ConectarBDMaterial()
        {
            listBoxMaterial.ClearSelected();
            listBoxMaterial.Items.Clear();
            ComunicaBDtekla BD = new ComunicaBDtekla();
            string Familia = "Painel";
            try
            {
                BD.ConectarBD();
                List<string> cores = BD.Procurarbd("SELECT [Material] FROM [dbo].[MateriaisSuplementares] WHERE [Familia]='" + Familia.Trim() + "'");

                listBoxMaterial.Items.Clear();
                foreach (string cor in cores)
                {
                    listBoxMaterial.Items.Add(cor);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Erro ao buscar dados: " + ex.Message);
            }
            finally
            {
                BD.DesonectarBD();
            }
        }

        private void InserirCor()
        {
            string novaCor = textBoxCor.Text.Trim();
            string familia = "Policarbonato"; // Ou use outro valor se necessário

            if (string.IsNullOrWhiteSpace(novaCor))
            {
                MessageBox.Show("Insira um valor para a cor.");
                return;
            }
            ComunicaBDtekla BD = new ComunicaBDtekla();
            try
            {
                BD.ConectarBD();
                string query = "INSERT INTO [dbo].[corpolicarbonato] ([Cor], [Familia]) VALUES (@cor, @familia)";
                using (SqlCommand cmd = new SqlCommand(query, BD.GetConnection()))
                {
                    cmd.Parameters.AddWithValue("@cor", novaCor);
                    cmd.Parameters.AddWithValue("@familia", familia);
                    cmd.ExecuteNonQuery();
                }

                MessageBox.Show("Cor adicionada com sucesso!");
                ConectarBDcor();
                textBoxCor.Clear();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Erro ao inserir cor: " + ex.Message);
            }
            finally
            {
                BD.DesonectarBD();
            }
        }

        private void RemoverCorSelecionada()
        {
            if (listBoxCor.SelectedItem == null)
            {
                MessageBox.Show("Por favor, selecione uma cor para remover.");
                return;
            }
            string corSelecionada = listBoxCor.SelectedItem.ToString();
            string familia = "Policarbonato";
            ComunicaBDtekla BD = new ComunicaBDtekla();
            try
            {
                BD.ConectarBD();

                string query = "DELETE FROM [dbo].[corpolicarbonato] WHERE [Cor] = @cor AND [Familia] = @familia";

                using (SqlCommand cmd = new SqlCommand(query, BD.GetConnection()))
                {
                    cmd.Parameters.AddWithValue("@cor", corSelecionada);
                    cmd.Parameters.AddWithValue("@familia", familia);
                    int rowsAffected = cmd.ExecuteNonQuery();

                    if (rowsAffected > 0)
                    {
                        MessageBox.Show("Cor removida com sucesso!");
                    }
                    else
                    {
                        MessageBox.Show("Nenhuma cor foi removida. Verifique a seleção.");
                    }
                }
                ConectarBDcor();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Erro ao remover cor: " + ex.Message);
            }
            finally
            {
                BD.DesonectarBD();
            }
        }

        private void InserirMaterial()
        {
            string Material = textBoxMaterial.Text.Trim();
            string familia = "Painel";
            if (string.IsNullOrWhiteSpace(Material))
            {
                MessageBox.Show("Insira um valor para o Material.");
                return;
            }
            ComunicaBDtekla BD = new ComunicaBDtekla();
            try
            {
                BD.ConectarBD();
                string query = "INSERT INTO [dbo].[MateriaisSuplementares] ([Material], [Familia]) VALUES (@Material, @familia)";
                using (SqlCommand cmd = new SqlCommand(query, BD.GetConnection()))
                {
                    cmd.Parameters.AddWithValue("@Material", Material);
                    cmd.Parameters.AddWithValue("@familia", familia);
                    cmd.ExecuteNonQuery();
                }

                MessageBox.Show("Cor adicionada com sucesso!");

                ConectarBDMaterial();
                textBoxMaterial.Clear();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Erro ao inserir cor: " + ex.Message);
            }
            finally
            {
                BD.DesonectarBD();
            }
        }

        private void RemoverMaterialSelecionada()
        {
            if (listBoxMaterial.SelectedItem == null)
            {
                MessageBox.Show("Por favor, selecione uma cor para remover.");
                return;
            }
            string Material = listBoxMaterial.SelectedItem.ToString();
            string familia = "Painel";
            ComunicaBDtekla BD = new ComunicaBDtekla();
            try
            {
                BD.ConectarBD();
                string query = "DELETE FROM [dbo].[MateriaisSuplementares] WHERE [Material] = @Material AND [Familia] = @familia";
                using (SqlCommand cmd = new SqlCommand(query, BD.GetConnection()))
                {
                    cmd.Parameters.AddWithValue("@Material", Material);
                    cmd.Parameters.AddWithValue("@familia", familia);
                    int rowsAffected = cmd.ExecuteNonQuery();

                    if (rowsAffected > 0)
                    {
                        MessageBox.Show("Material removido com sucesso!");
                    }
                    else
                    {
                        MessageBox.Show("Nenhum Material foi removido. Verifique se esta selecionado.");
                    }
                }

                ConectarBDMaterial();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Erro ao remover cor: " + ex.Message);
            }
            finally
            {
                BD.DesonectarBD();
            }
        }
    }
}

