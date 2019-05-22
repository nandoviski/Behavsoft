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
		XmlDocument docMenu;
		string[] listaExcel;
		
		public ucJuntarExcel(string caminhoXml)
		{
			InitializeComponent();
			IniciarCombo(caminhoXml);
		}

		void IniciarCombo(string caminhoXml)
		{
			// Limpar itens ComboBox
			if (cbTipoComportamento.Items.Count > 0)
			{
				var itensCombo = cbTipoComportamento.Items.Count;
				for (var i = itensCombo - 1; i >= 0; i--)
				{
					cbTipoComportamento.Items.RemoveAt(i);
				}
			}

			docMenu = new XmlDocument();
			docMenu.Load(caminhoXml);

			var menus = docMenu.SelectNodes("Menus");

			if (menus != null && menus.Count > 0)
			{
				foreach (XmlNode item in menus[0].ChildNodes)
				{
					var itemCombo = new ComboProtocoloItem();
					itemCombo.Nome = item.Attributes["nome"].Value;
					itemCombo.Nodos = item.ChildNodes;

					cbTipoComportamento.Items.Add(itemCombo);
				}
			}

			cbTipoComportamento.SelectedIndex = 0;
		}

		void cbTipoComportamento_SelectionChanged(object sender, SelectionChangedEventArgs e)
		{
			if (cbTipoComportamento.SelectedItem != null)
			{
				var comboItem = (ComboProtocoloItem)cbTipoComportamento.SelectedItem;

				if (comboItem.Nodos?.Count > 0)
				{
					foreach (XmlNode item in comboItem.Nodos)
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

		void btnBuscarExcels_Click(object sender, RoutedEventArgs e)
		{
			var oSelWc = new OpenFileDialog();
			oSelWc.Multiselect = true;
			oSelWc.Filter = "Excel (*.xls)|*.xls";
			oSelWc.Title = "Excel Files";
			oSelWc.CheckFileExists = true;
			oSelWc.CheckPathExists = true;

			var result = oSelWc.ShowDialog();

			if (result.HasValue && result.Value)
			{
				lbAcao.Items.Clear();
				listaExcel = oSelWc.FileNames;

				foreach (string sArq in oSelWc.FileNames)
					lbAcao.Items.Add(sArq);
			}
		}

		void btnGerarExcel_Click(object sender, RoutedEventArgs e)
		{
			if (string.IsNullOrEmpty(txtSalvarEm.Text))
			{
				MessageBox.Show("Please informe the excel save path", "Warning", MessageBoxButton.OK, MessageBoxImage.Warning);
				return;
			}

			if (listaExcel == null || listaExcel.Length < 1)
			{
				MessageBox.Show("Please informe the files that will be merged.", "Warning", MessageBoxButton.OK, MessageBoxImage.Warning);
				return;
			}

			if (System.IO.File.Exists(txtSalvarEm.Text))
			{
				var result = MessageBox.Show("File '" + txtSalvarEm.Text + "' already exists.\nDo you want to replace it?", "Information", MessageBoxButton.YesNo, MessageBoxImage.Information);
				if (result == MessageBoxResult.No)
					return;
			}

			var nomeTecla = new Dictionary<string, string>();
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

		void btnSalvarEm_Click(object sender, RoutedEventArgs e)
		{
			var sabeDlg = new SaveFileDialog();
			sabeDlg.FileName = txtSalvarEm.Text;
			sabeDlg.Filter = "Excel 97-2003 Workbook|*.xls|Excel Workbook|*.xlsx";

			var ret = sabeDlg.ShowDialog();

			if (ret.HasValue && ret.Value)
			{
				txtSalvarEm.Text = sabeDlg.FileName;
			}
		}
	}
}
