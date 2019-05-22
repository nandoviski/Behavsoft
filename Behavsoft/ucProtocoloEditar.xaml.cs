using System;
using System.Collections.Generic;
using System.Linq;
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
	/// Interaction logic for ucProtocoloEditar.xaml
	/// </summary>
	public partial class ucProtocoloEditar : Window
	{
		string caminhoXml;
		XmlDocument docMenu;
		int ultimoItemSelecionado = -1;

		public ucProtocoloEditar(string caminhoXml)
		{
			InitializeComponent();

			this.caminhoXml = caminhoXml;
			this.IniciarCombo(caminhoXml);
		}

		void IniciarCombo(string caminhoXml)
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

			docMenu = new XmlDocument();
			docMenu.Load(caminhoXml);

			XmlNodeList menus = docMenu.SelectNodes("Menus");

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

			if (ultimoItemSelecionado >= 0)
			{
				if (cbTipoComportamento.Items.Count > ultimoItemSelecionado)
					cbTipoComportamento.SelectedIndex = ultimoItemSelecionado;
			}
			else
			{
				if (cbTipoComportamento.Items.Count > 0)
					cbTipoComportamento.SelectedIndex = 0;
			}
		}

		void cbTipoComportamento_SelectionChanged(object sender, SelectionChangedEventArgs e)
		{
			if (cbTipoComportamento.SelectedItem != null)
			{
				var comboItem = (ComboProtocoloItem)cbTipoComportamento.SelectedItem;
				ultimoItemSelecionado = cbTipoComportamento.SelectedIndex;

				if (comboItem.Nodos != null && comboItem.Nodos.Count > 0)
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

		void btnEditar_Click(object sender, RoutedEventArgs e)
		{
			this.HabilitarEdicao(true);
		}

		void btnSalvar_Click(object sender, RoutedEventArgs e)
		{
			XmlNodeList menus = docMenu.SelectNodes("Menus");

			if (menus != null && menus.Count > 0)
			{
				var itemCombo = (ComboProtocoloItem)cbTipoComportamento.SelectedItem;

				foreach (XmlNode item in menus[0].ChildNodes)
				{
					if (item.Attributes["nome"].Value == itemCombo.Nome)
					{
						foreach (XmlNode nodo in item.ChildNodes)
						{
							if (nodo.Attributes["tecla"].Value == "A")
								nodo.InnerText = txtTeclaA.Text;
							else if (nodo.Attributes["tecla"].Value == "S")
								nodo.InnerText = txtTeclaS.Text;
							else if (nodo.Attributes["tecla"].Value == "D")
								nodo.InnerText = txtTeclaD.Text;
							else if (nodo.Attributes["tecla"].Value == "F")
								nodo.InnerText = txtTeclaF.Text;
							else if (nodo.Attributes["tecla"].Value == "H")
								nodo.InnerText = txtTeclaH.Text;
							else if (nodo.Attributes["tecla"].Value == "J")
								nodo.InnerText = txtTeclaJ.Text;
							else if (nodo.Attributes["tecla"].Value == "K")
								nodo.InnerText = txtTeclaK.Text;
							else if (nodo.Attributes["tecla"].Value == "L")
								nodo.InnerText = txtTeclaL.Text;
						}
						docMenu.Save(caminhoXml);
						IniciarCombo(caminhoXml);
						break;
					}
				}
			}

			this.HabilitarEdicao(false);
		}

		void btnCancelar_Click(object sender, RoutedEventArgs e)
		{
			HabilitarEdicao(false);
		}

		void HabilitarEdicao(bool habilitar)
		{
			gbTeclas.IsEnabled = habilitar;
			btnSalvar.IsEnabled = habilitar;
			btnCancelar.IsEnabled = habilitar;

			cbTipoComportamento.IsEnabled = !habilitar;
			btnEditar.IsEnabled = !habilitar;
		}
	}
}
