using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;

namespace Behavsoft
{
    /// <summary>
    /// Interaction logic for JuntarExcel.xaml
    /// </summary>
    public partial class JuntarExcel : Window
    {
        public JuntarExcel()
        {
            InitializeComponent();

            cbTipoComportamento.SelectionChanged += cbTipoComportamento_SelectionChanged;
            cbTipoComportamento.Items.Add("Labirinto Em Cruz Elevado");
            cbTipoComportamento.Items.Add("Campo Aberto");
            cbTipoComportamento.Items.Add("Nado Forçado");
            cbTipoComportamento.SelectedIndex = 0;

            cbTipoExcel.Items.Add("Novo");
            cbTipoExcel.Items.Add("Antigo");
            cbTipoExcel.SelectedIndex = 0;
        }

        private void cbTipoComportamento_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            // Nao mudar Index
            if (cbTipoComportamento.SelectedIndex == 0)
            {
                txtTeclaA.Text = "Braço Aberto";
                txtTeclaS.Text = "";
                txtTeclaD.Text = "Grooming";
                txtTeclaF.Text = "Braço Fechado";
                txtTeclaH.Text = "";
                txtTeclaJ.Text = "Esticar";
                txtTeclaK.Text = "Espiar";
                txtTeclaL.Text = "";
            }
            else if (cbTipoComportamento.SelectedIndex == 1)
            {
                txtTeclaA.Text = "Andar";
                txtTeclaS.Text = "Levantar";
                txtTeclaD.Text = "Grooming";
                txtTeclaF.Text = "Fezes";
                txtTeclaH.Text = "";
                txtTeclaJ.Text = "Periferia";
                txtTeclaK.Text = "Central";
                txtTeclaL.Text = "";
            }
            else if (cbTipoComportamento.SelectedIndex == 2)
            {
                txtTeclaA.Text = "Escalar";
                txtTeclaS.Text = "Nadar";
                txtTeclaD.Text = "Headshake";
                txtTeclaF.Text = "";
                txtTeclaH.Text = "";
                txtTeclaJ.Text = "Flutuar";
                txtTeclaK.Text = "Congelar";
                txtTeclaL.Text = "Mergulhar";
            }
        }


        private string[] listaExcel;
        private string caminhoExcel = string.Empty;
        private void btnBuscarExcels_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog oSelWc = new OpenFileDialog();

            oSelWc.Multiselect = true;
            oSelWc.Filter = "Excel (*.xls)|*.xls";
            oSelWc.Title = "Arquivos Excel";
            oSelWc.CheckFileExists = true;
            oSelWc.CheckPathExists = true;

            if (string.IsNullOrEmpty(caminhoExcel))
                caminhoExcel = AppDomain.CurrentDomain.BaseDirectory;

            oSelWc.InitialDirectory = caminhoExcel;

            bool? result = oSelWc.ShowDialog();

            if (result.HasValue && result.Value)
            {
                lbAcao.Items.Clear();
                listaExcel = oSelWc.FileNames;

                System.IO.FileInfo a = new System.IO.FileInfo(oSelWc.FileName);
                if(a != null)
                    caminhoExcel = a.Directory.ToString();
                
                foreach (String sArq in oSelWc.FileNames)
                    lbAcao.Items.Add(sArq);
            }
        }

        private void btnGerarExcel_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrEmpty(txtSalvarEm.Text))
            {
                MessageBox.Show("É obrigatório informar um arquivo a ser salvo", "Atenção", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            if (listaExcel == null || listaExcel.Length < 1)
            {
                MessageBox.Show("É obrigatório selecionar algum arquivo excel.", "Atenção", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            if (System.IO.File.Exists(txtSalvarEm.Text))
            {
                MessageBoxResult result = MessageBox.Show("O arquivo '" + txtSalvarEm.Text + "' já existe.\nDeseja substituí-lo?", "Atenção", MessageBoxButton.YesNo, MessageBoxImage.Information);
                if (result == MessageBoxResult.No)
                    return;
            }

            Dictionary<string, string> nomeTecla = new Dictionary<string, string>();
            nomeTecla.Add("A", txtTeclaA.Text);
            nomeTecla.Add("S", txtTeclaS.Text);
            nomeTecla.Add("D", txtTeclaD.Text);
            nomeTecla.Add("F", txtTeclaF.Text);
            nomeTecla.Add("H", txtTeclaH.Text);
            nomeTecla.Add("J", txtTeclaJ.Text);
            nomeTecla.Add("K", txtTeclaK.Text);
            nomeTecla.Add("L", txtTeclaL.Text);

            new ExcelUtil().MesclarExcel(listaExcel, nomeTecla, txtSalvarEm.Text, cbTipoExcel.SelectedItem.ToString());
        }

        private void btnSalvarEm_Click(object sender, RoutedEventArgs e)
        {
            Microsoft.Win32.SaveFileDialog sabeDlg = new Microsoft.Win32.SaveFileDialog();
            sabeDlg.FileName = txtSalvarEm.Text;
            
            sabeDlg.Filter = "Pasta de Trabalho do Excel 97-2003|*.xls|Pasta de Trabalho do Excel|*.xlsx";
            bool? ret = sabeDlg.ShowDialog();

            if (ret.HasValue && ret.Value)
            {
                txtSalvarEm.Text = sabeDlg.FileName;
            }

        }

        
    }
}
