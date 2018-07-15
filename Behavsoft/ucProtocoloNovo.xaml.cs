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
    /// Interaction logic for ucProtocoloNovo.xaml
    /// </summary>
    public partial class ucProtocoloNovo : Window
    {
        private string _caminhoXml;
        private XmlDocument _docMenu;

        public ucProtocoloNovo(string caminhoXml)
        {
            InitializeComponent();

            _caminhoXml = caminhoXml;
            this.IniciarCombo(caminhoXml);
        }

        private void IniciarCombo(string caminhoXml)
        {
            _docMenu = new XmlDocument();
            _docMenu.Load(caminhoXml);

            //XmlNodeList menus = _docMenu.SelectNodes("Menus");

            //if (menus != null && menus.Count > 0)
            //{
            //    //foreach (XmlNode item in menus[0].ChildNodes)
            //    //{
            //    //    ComboProtocoloItem itemCombo = new ComboProtocoloItem();
            //    //    itemCombo.Nome = item.Attributes["nome"].Value;
            //    //    itemCombo.Nodos = item.ChildNodes;

            //    //    cbTipoComportamento.Items.Add(itemCombo);
            //    //}
            //}

        }

        private void btnSalvar_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrEmpty(tbNomeProtocolo.Text))
            {
                MessageBox.Show("O nome do protocolo é obrigatório", "Atenção", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            XmlNodeList menus = _docMenu.SelectNodes("Menus");

            if (menus != null && menus.Count > 0)
            {
                // Valida se o nome já existe
                foreach (XmlNode item in menus[0].ChildNodes)
                {
                    if (item.Attributes["nome"].Value.ToUpper() == tbNomeProtocolo.Text.ToUpper())
                    {
                        MessageBox.Show("O nome do protocolo informado já existe", "Atenção", MessageBoxButton.OK, MessageBoxImage.Warning);
                        return;
                    }
                }

                XmlElement elemProtocolo = _docMenu.CreateElement("protocolo");
                elemProtocolo.SetAttribute("nome", tbNomeProtocolo.Text);

                // Tecla A
                XmlElement elemTeclaA = _docMenu.CreateElement("item");
                elemTeclaA.SetAttribute("tecla", "A");
                elemTeclaA.InnerText = txtTeclaA.Text;
                elemProtocolo.AppendChild(elemTeclaA);

                // Tecla S
                XmlElement elemTeclaS = _docMenu.CreateElement("item");
                elemTeclaS.SetAttribute("tecla", "S");
                elemTeclaS.InnerText = txtTeclaS.Text;
                elemProtocolo.AppendChild(elemTeclaS);

                // Tecla D
                XmlElement elemTeclaD = _docMenu.CreateElement("item");
                elemTeclaD.SetAttribute("tecla", "D");
                elemTeclaD.InnerText = txtTeclaD.Text;
                elemProtocolo.AppendChild(elemTeclaD);

                // Tecla F
                XmlElement elemTeclaF = _docMenu.CreateElement("item");
                elemTeclaF.SetAttribute("tecla", "F");
                elemTeclaF.InnerText = txtTeclaF.Text;
                elemProtocolo.AppendChild(elemTeclaF);

                // Tecla H
                XmlElement elemTeclaH = _docMenu.CreateElement("item");
                elemTeclaH.SetAttribute("tecla", "H");
                elemTeclaH.InnerText = txtTeclaH.Text;
                elemProtocolo.AppendChild(elemTeclaH);

                // Tecla J
                XmlElement elemTeclaJ = _docMenu.CreateElement("item");
                elemTeclaJ.SetAttribute("tecla", "J");
                elemTeclaJ.InnerText = txtTeclaJ.Text;
                elemProtocolo.AppendChild(elemTeclaJ);

                // Tecla K
                XmlElement elemTeclaK = _docMenu.CreateElement("item");
                elemTeclaK.SetAttribute("tecla", "K");
                elemTeclaK.InnerText = txtTeclaK.Text;
                elemProtocolo.AppendChild(elemTeclaK);

                // Tecla L
                XmlElement elemTeclaL = _docMenu.CreateElement("item");
                elemTeclaL.SetAttribute("tecla", "L");
                elemTeclaL.InnerText = txtTeclaL.Text;
                elemProtocolo.AppendChild(elemTeclaL);

                menus[0].AppendChild(elemProtocolo);

                _docMenu.Save(_caminhoXml);

                MessageBox.Show("Novo protocolo criado com sucesso", "Sucesso", MessageBoxButton.OK, MessageBoxImage.Information);
                this.Close();
            }
        }

        private void btnCancelar_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }
    }
}
