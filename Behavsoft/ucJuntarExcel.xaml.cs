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
using System.Xml;

namespace Behavsoft
{
	/// <summary>
	/// Interaction logic for ucJuntarExcel.xaml
	/// </summary>
	public partial class ucJuntarExcel : Window
	{
		private XmlDocument _docMenu;
		private string[] listaExcel;
		private string caminhoExcel = string.Empty;

		public ucJuntarExcel(string caminhoXml)
		{
			InitializeComponent();

			this.IniciarCombo(caminhoXml);
		}

		private void IniciarCombo(string caminhoXml)
		{
			// Limpar itens ComboBox
			if (cbTipoComportamento.Items.Count > 0)
			{
				int itensCombo = cbTipoComportamento.Items.Count;
				for (int i = itensCombo - 1; i >= 0; i--)
				{
					cbTipoComportamento.Items.RemoveAt(i);
				}
			}

			_docMenu = new XmlDocument();
			_docMenu.Load(caminhoXml);

			XmlNodeList menus = _docMenu.SelectNodes("Menus");

			if (menus != null && menus.Count > 0)
			{
				foreach (XmlNode item in menus[0].ChildNodes)
				{
					ComboProtocoloItem itemCombo = new ComboProtocoloItem();
					itemCombo.Nome = item.Attributes["nome"].Value;
					itemCombo.Nodos = item.ChildNodes;

					cbTipoComportamento.Items.Add(itemCombo);
				}
			}

			cbTipoComportamento.SelectedIndex = 0;
		}

		private void cbTipoComportamento_SelectionChanged(object sender, SelectionChangedEventArgs e)
		{
			if (cbTipoComportamento.SelectedItem != null)
			{
				ComboProtocoloItem teste = (ComboProtocoloItem)cbTipoComportamento.SelectedItem;

				if (teste.Nodos != null && teste.Nodos.Count > 0)
				{
					foreach (XmlNode item in teste.Nodos)
					{
						if (item.Attributes["tecla"].Value == "A")
							txtTeclaA.Text = item.InnerText;
						else if (item.Attributes["tecla"].Value == "S")
							txtTeclaS.Text = item.InnerText;
						else if (item.Attributes["tecla"].Value == "D")
							txtTeclaD.Text = item.InnerText;
						else if (item.Attributes["tecla"].Value == "F")
							txtTeclaF.Text = item.InnerText;
						else if (item.Attributes["tecla"].Value == "H")
							txtTeclaH.Text = item.InnerText;
						else if (item.Attributes["tecla"].Value == "J")
							txtTeclaJ.Text = item.InnerText;
						else if (item.Attributes["tecla"].Value == "K")
							txtTeclaK.Text = item.InnerText;
						else if (item.Attributes["tecla"].Value == "L")
							txtTeclaL.Text = item.InnerText;
					}
				}
			}
		}

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
				if (a != null)
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

			new ExcelUtil().MesclarExcel(listaExcel, nomeTecla, txtSalvarEm.Text);
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
