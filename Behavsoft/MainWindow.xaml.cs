using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Threading;
using System.Xml;

namespace Behavsoft
{
	/// <summary>
	/// Interaction logic for MainWindow.xaml
	/// </summary>
	public partial class MainWindow : Window
	{
		TimeSpan? ultimoClique;
		Key ultimaTecla;
		List<TemposItem> tempos;
		VideoState estadoVideo;
		DispatcherTimer timer;
		bool isDragging;
		TimeSpan? inicioAnalise;
		int ultimoItemSelecionado = -1;

		string _caminhoLocal
		{
			get
			{
				return AppDomain.CurrentDomain.BaseDirectory;
			}
		}

		string caminhoXml
		{
			get
			{
				return _caminhoLocal + "MenuConfiguration.xml";
			}
		}

		public MainWindow()
		{
			InitializeComponent();

			KeyDown += MainWindow_KeyDown;
			this.WindowState = System.Windows.WindowState.Maximized;

			tempos = new List<TemposItem>();
			latencia = new Dictionary<Key, TimeSpan>();
			estadoVideo = VideoState.Stopped;

			cbTipoComportamento.SelectionChanged += cbTipoComportamento_SelectionChanged;
			cbTipoComportamento.DisplayMemberPath = "Nome";
			this.IniciarProtocolos();

			inicioAnalise = null;

			isDragging = false;
			timer = new DispatcherTimer();
			timer.Interval = TimeSpan.FromMilliseconds(300);
			timer.Tick += timer_Tick;
		}

		void timer_Tick(object sender, EventArgs e)
		{
			if (!isDragging)
				sldBarraTempo.Value = VideoControl.Position.TotalSeconds;

			if (inicioAnalise.HasValue)
			{
				TimeSpan tempoAnalise = VideoControl.Position - inicioAnalise.Value;
				if (tempoAnalise.TotalSeconds > 1)
				{
					txtTempoAnalise.Text = tempoAnalise.Hours.ToString("00") + ":" + tempoAnalise.Minutes.ToString("00") + ":" + tempoAnalise.Seconds.ToString("00");

					if (tempoAnalise.Minutes >= 5 || (tempoAnalise.Minutes >= 4 && tempoAnalise.Seconds >= 45))
					{
						if (txtTempoAnalise.Foreground is SolidColorBrush &&
							(txtTempoAnalise.Foreground as SolidColorBrush).Color == Colors.Black)
							txtTempoAnalise.Foreground = new SolidColorBrush(Colors.Red);
						else
							txtTempoAnalise.Foreground = new SolidColorBrush(Colors.Black);
						txtTempoAnalise.FontWeight = FontWeights.Bold;
					}
				}
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

		void MainWindow_KeyDown(object sender, KeyEventArgs e)
		{
			if (estadoVideo == VideoState.Played)
			{
				if ((e.Key == Key.A && !string.IsNullOrEmpty(txtTeclaA.Text)) ||
					(e.Key == Key.S && !string.IsNullOrEmpty(txtTeclaS.Text)) ||
					(e.Key == Key.D && !string.IsNullOrEmpty(txtTeclaD.Text)) ||
					(e.Key == Key.F && !string.IsNullOrEmpty(txtTeclaF.Text)) ||
					(e.Key == Key.H && !string.IsNullOrEmpty(txtTeclaH.Text)) ||
					(e.Key == Key.J && !string.IsNullOrEmpty(txtTeclaJ.Text)) ||
					(e.Key == Key.K && !string.IsNullOrEmpty(txtTeclaK.Text)) ||
					(e.Key == Key.L && !string.IsNullOrEmpty(txtTeclaL.Text)))
				{
					var tempoAgora = new TimeSpan(VideoControl.Position.Hours, VideoControl.Position.Minutes, VideoControl.Position.Seconds);

					if (!inicioAnalise.HasValue)
					{
						txtTempoAnalise.Foreground = new SolidColorBrush(Colors.Black);
						txtTempoAnalise.FontWeight = FontWeights.Normal;
						txtTempoAnalise.Text = "00:00:00";
						inicioAnalise = VideoControl.Position;
					}

					//if (!ultimoClique.HasValue)
					//{
					//    lbAcao.Items.Add("Iniciou");
					//}
					//else
					{
						var tempo = BuscarUltimoTempo(e.Key);
						if (tempo == null)
						{
							tempos.Add(new TemposItem() { Tecla = e.Key, Inicio = tempoAgora, TextoAtalho = BuscarTextoAtalho(e.Key) });
							this.MudarCorTecla(e.Key, true);
						}
						else
						{
							tempo.Fim = tempoAgora;
							this.MudarCorTecla(e.Key, false);
						}
						Latencia(e.Key, tempoAgora);
					}

					ultimoClique = tempoAgora;
					ultimaTecla = e.Key;

					MostrarTempos();
				}
			}

			if (e.Key == Key.P)
			{
				if (txtTeclaA.IsFocused || txtTeclaS.IsFocused || txtTeclaD.IsFocused || txtTeclaF.IsFocused ||
					txtTeclaH.IsFocused || txtTeclaJ.IsFocused || txtTeclaK.IsFocused || txtTeclaL.IsFocused)
					return;

				if (estadoVideo == VideoState.Paused || estadoVideo == VideoState.Stopped)
				{
					PlayButton_Click_1(null, null);
				}
				else if (estadoVideo == VideoState.Played)
				{
					PauseButton_Click_1(null, null);
				}
			}

			if (e.Key == Key.Escape)
			{
				if (estadoVideo == VideoState.Paused || estadoVideo == VideoState.Played)
				{
					StopButton_Click_1(null, null);
				}
			}

			if (e.Key == Key.S && (Keyboard.Modifiers & ModifierKeys.Control) == ModifierKeys.Control)
			{
				btnSalvarComo_Click(null, null);
			}

			if (e.Key == Key.N && (Keyboard.Modifiers & ModifierKeys.Control) == ModifierKeys.Control)
			{
				ucJuntarExcel j = new ucJuntarExcel(caminhoXml);
				j.ShowDialog();
			}
		}

		TemposItem BuscarUltimoTempo(Key tecla)
		{
			if (tempos != null && tempos.Count > 0)
			{
				foreach (TemposItem item in tempos)
				{
					if (item.Tecla == tecla && !item.Fim.HasValue)
						return item;
				}
			}

			return null;
		}

		Dictionary<Key, TimeSpan> latencia = new Dictionary<Key, TimeSpan>();
		void Latencia(Key tecla, TimeSpan tempo)
		{
			if (!latencia.ContainsKey(tecla))
			{
				latencia.Add(tecla, tempo);
			}
		}

		void BrowseButton_Click_1(object sender, RoutedEventArgs e)
		{
			try
			{
				var openDlg = new Microsoft.Win32.OpenFileDialog();
				var ret = openDlg.ShowDialog();

				if (ret.HasValue && ret.Value)
				{
					MediaPathTextBox.Text = openDlg.FileName;
					VideoControl.Source = new Uri(openDlg.FileName);
					estadoVideo = VideoState.Stopped;
					tempos = new List<TemposItem>();
					latencia = new Dictionary<Key, TimeSpan>();

					VideoControl.Play();
					for (int i = 0; i < 2500000; i++)
					{
						// Espera tempo para mostrar o inicio do video
					}
					VideoControl.Stop();
				}
			}
			catch
			{
				MessageBox.Show("Error loading the video", "Erro", MessageBoxButton.OK, MessageBoxImage.Error);
			}
		}

		void PlayButton_Click_1(object sender, RoutedEventArgs e)
		{
			if (MediaPathTextBox.Text.Length <= 0)
			{
				MessageBox.Show("Please select a video", "Warning", MessageBoxButton.OK, MessageBoxImage.Warning);
				return;
			}

			if (estadoVideo == VideoState.Stopped)
			{
				txtTempoAnalise.Foreground = new SolidColorBrush(Colors.Black);
				txtTempoAnalise.FontWeight = FontWeights.Normal;
				txtTempoAnalise.Text = "00:00:00";
				tempos = new List<TemposItem>();
				latencia = new Dictionary<Key, TimeSpan>();
				MostrarTempos();
			}

			try
			{
				VideoControl.Play();
				estadoVideo = VideoState.Played;
				BloquearComponentesConfig(true);

				while (!VideoControl.NaturalDuration.HasTimeSpan)
				{
				}

				TimeSpan ts = VideoControl.NaturalDuration.TimeSpan;
				sldBarraTempo.Maximum = ts.TotalSeconds;
				sldBarraTempo.SmallChange = 1;
				sldBarraTempo.LargeChange = Math.Min(10, ts.Seconds / 10);

				timer.Start();
			}
			catch (Exception)
			{
			}
		}

		void PauseButton_Click_1(object sender, RoutedEventArgs e)
		{
			VideoControl.Pause();
			estadoVideo = VideoState.Paused;
			timer.Stop();
		}

		void StopButton_Click_1(object sender, RoutedEventArgs e)
		{
			ultimoClique = null;
			FinalizarTodasTeclas(VideoControl.Position);
			MostrarTempos();
			VideoControl.Stop();

			estadoVideo = VideoState.Stopped;
			BloquearComponentesConfig(false);

			timer.Stop();
			inicioAnalise = null;
		}

		void MostrarTempos()
		{
			lbAcao.Items.Clear();

			foreach (var item in tempos)
			{
				var msg = item.Tecla + " " + item.Inicio.Value;//.ToString(@"hh\:mm\:ss");

				if (item.Fim.HasValue)
				{
					msg += " " + item.Fim.Value + " " + item.Duracao(); //.ToString(@"hh\:mm\:ss")
				}

				lbAcao.Items.Add(msg);
			}

			if (lbAcao.Items.Count > 0)
				lbAcao.ScrollIntoView(lbAcao.Items[lbAcao.Items.Count - 1]);
		}

		void BloquearComponentesConfig(bool bloquear)
		{
			BrowseButton.IsEnabled = !bloquear;
			gbTeclas.IsEnabled = !bloquear;
			cbTipoComportamento.IsEnabled = !bloquear;
		}

		void btnSalvarComo_Click(object sender, RoutedEventArgs e)
		{
			if (estadoVideo != VideoState.Stopped)
			{
				MessageBox.Show("Please stop the video before save", "Warning", MessageBoxButton.OK, MessageBoxImage.Warning);
				return;
			}

			if (tempos == null || tempos.Count < 1)
			{
				MessageBox.Show("There is no data to be saved", "Warning", MessageBoxButton.OK, MessageBoxImage.Warning);
				return;
			}

			var sabeDlg = new Microsoft.Win32.SaveFileDialog();
			sabeDlg.InitialDirectory = @"c:\";
			sabeDlg.Filter = "Excel 97-2003 Workbook|*.xls|Excel Workbook|*.xlsx";
			var ret = sabeDlg.ShowDialog();

			if (ret.HasValue && ret.Value)
			{
				var util = new ExcelUtil();
				util.JanelaPai = this;
				util.GerarExcel(sabeDlg.FileName, tempos);

				//EPPlusUtil a = new EPPlusUtil();
				//a.GerarExcelTESTE(sabeDlg.FileName, tempos);
			}
		}

		void sldBarraTempo_DragStarted_1(object sender, System.Windows.Controls.Primitives.DragStartedEventArgs e)
		{
			isDragging = true;
		}

		void sldBarraTempo_ValueChanged(object sender, RoutedPropertyChangedEventArgs<double> e)
		{
			if (isDragging)
				VideoControl.Position = TimeSpan.FromSeconds(sldBarraTempo.Value);

			if (VideoControl.NaturalDuration.HasTimeSpan)
			{
				txtTempoVideo.Text =
				   VideoControl.Position.Hours.ToString("00") + ":" + VideoControl.Position.Minutes.ToString("00") + ":" + VideoControl.Position.Seconds.ToString("00") +
				   " / " +
				   VideoControl.NaturalDuration.TimeSpan.Hours.ToString("00") + ":" + VideoControl.NaturalDuration.TimeSpan.Minutes.ToString("00") + ":" + VideoControl.NaturalDuration.TimeSpan.Seconds.ToString("00");
			}
		}

		void sldBarraTempo_DragCompleted_1(object sender, System.Windows.Controls.Primitives.DragCompletedEventArgs e)
		{
			isDragging = false;
			//VideoControl.Position = TimeSpan.FromSeconds(sldBarraTempo.Value);
		}

		// Finaliza todas as teclas aberta a o clicar no Stop
		void FinalizarTodasTeclas(TimeSpan final)
		{
			var tempo = new TimeSpan(final.Hours, final.Minutes, final.Seconds);

			foreach (TemposItem item in tempos)
			{
				this.MudarCorTecla(item.Tecla, false);
				if (!item.Fim.HasValue)
					item.Fim = tempo;
			}
		}

		string BuscarTextoAtalho(Key tecla)
		{
			string retorno = string.Empty;

			switch (tecla)
			{
				case Key.A:
					retorno = txtTeclaA.Text;
					break;
				case Key.S:
					retorno = txtTeclaS.Text;
					break;
				case Key.D:
					retorno = txtTeclaD.Text;
					break;
				case Key.F:
					retorno = txtTeclaF.Text;
					break;
				case Key.H:
					retorno = txtTeclaH.Text;
					break;
				case Key.J:
					retorno = txtTeclaJ.Text;
					break;
				case Key.K:
					retorno = txtTeclaK.Text;
					break;
				case Key.L:
					retorno = txtTeclaL.Text;
					break;
			}

			return retorno;
		}

		void MudarCorTecla(Key tecla, bool isPressionado)
		{
			switch (tecla)
			{
				case Key.A:
					if (isPressionado)
						tbTeclaA.Style = (Style)FindResource("TeclaAtalhoPressionada");
					else
						tbTeclaA.Style = (Style)FindResource("TeclaAtalhoNormal");
					break;
				case Key.S:
					if (isPressionado)
						tbTeclaS.Style = (Style)FindResource("TeclaAtalhoPressionada");
					else
						tbTeclaS.Style = (Style)FindResource("TeclaAtalhoNormal");
					break;
				case Key.D:
					if (isPressionado)
						tbTeclaD.Style = (Style)FindResource("TeclaAtalhoPressionada");
					else
						tbTeclaD.Style = (Style)FindResource("TeclaAtalhoNormal");
					break;
				case Key.F:
					if (isPressionado)
						tbTeclaF.Style = (Style)FindResource("TeclaAtalhoPressionada");
					else
						tbTeclaF.Style = (Style)FindResource("TeclaAtalhoNormal");
					break;
				case Key.H:
					if (isPressionado)
						tbTeclaH.Style = (Style)FindResource("TeclaAtalhoPressionada");
					else
						tbTeclaH.Style = (Style)FindResource("TeclaAtalhoNormal");
					break;
				case Key.J:
					if (isPressionado)
						tbTeclaJ.Style = (Style)FindResource("TeclaAtalhoPressionada");
					else
						tbTeclaJ.Style = (Style)FindResource("TeclaAtalhoNormal");
					break;
				case Key.K:
					if (isPressionado)
						tbTeclaK.Style = (Style)FindResource("TeclaAtalhoPressionada");
					else
						tbTeclaK.Style = (Style)FindResource("TeclaAtalhoNormal");
					break;
				case Key.L:
					if (isPressionado)
						tbTeclaL.Style = (Style)FindResource("TeclaAtalhoPressionada");
					else
						tbTeclaL.Style = (Style)FindResource("TeclaAtalhoNormal");
					break;
			}
		}

		public void MostrarUcCarregando(bool mostrar = true)
		{
			if (mostrar)
			{
				ucAguardeExcel.Visibility = System.Windows.Visibility.Visible;
				btnSalvarComo.IsEnabled = false;
				StopButton.IsEnabled = false;
				PauseButton.IsEnabled = false;
				PlayButton.IsEnabled = false;
				BrowseButton.IsEnabled = false;
			}
			else
			{
				ucAguardeExcel.Visibility = System.Windows.Visibility.Collapsed;
				btnSalvarComo.IsEnabled = true;
				StopButton.IsEnabled = true;
				PauseButton.IsEnabled = true;
				PlayButton.IsEnabled = true;
				BrowseButton.IsEnabled = true;
			}
		}

		void miSair_Click(object sender, RoutedEventArgs e)
		{
			this.Close();
		}

		void miJuntarExcel_Click(object sender, RoutedEventArgs e)
		{
			ucJuntarExcel j = new ucJuntarExcel(caminhoXml);
			j.ShowDialog();
		}

		void miNovoProtocolo_Click(object sender, RoutedEventArgs e)
		{
			ucProtocoloNovo pn = new ucProtocoloNovo(caminhoXml);
			pn.ShowDialog();
			this.IniciarProtocolos();
		}

		void miEditarProtocolo_Click(object sender, RoutedEventArgs e)
		{
			if (cbTipoComportamento != null && cbTipoComportamento.Items.Count > 0)
			{
				var pe = new ucProtocoloEditar(caminhoXml);
				pe.ShowDialog();
				this.IniciarProtocolos();
			}
		}

		void IniciarProtocolos()
		{
			if (!File.Exists(caminhoXml))
			{
				// Cria menus padrões
				XmlTextWriter xtw = new XmlTextWriter(caminhoXml, Encoding.UTF8);
				xtw.Formatting = Formatting.Indented;
				xtw.Indentation = 2;

				xtw.WriteStartDocument();
				xtw.WriteStartElement("Menus");

				// Monta todos os itens padroes
				ProtocoloPadrao(ref xtw);

				xtw.WriteEndElement();
				xtw.WriteEndDocument();

				xtw.Flush();
				xtw.Close();
			}

			this.LimparComboBoxProtocolo();

			XmlDocument doc = new XmlDocument();
			doc.Load(caminhoXml);

			XmlNodeList menus = doc.SelectNodes("Menus");

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

		void ProtocoloPadrao(ref XmlTextWriter xtw)
		{
			xtw.WriteStartElement("protocolo"); // proto ini
			xtw.WriteAttributeString("nome", "Elevated plus maze");
			this.CriaItemXml(ref xtw, "A", "Open arm");
			this.CriaItemXml(ref xtw, "S", "Central crossed");
			this.CriaItemXml(ref xtw, "D", "Grooming");
			this.CriaItemXml(ref xtw, "F", "Closed arm");
			this.CriaItemXml(ref xtw, "H", string.Empty);
			this.CriaItemXml(ref xtw, "J", "Strech-attend");
			this.CriaItemXml(ref xtw, "K", "Head dipping");
			this.CriaItemXml(ref xtw, "L", "Rearing");
			xtw.WriteEndElement(); // proto fim

			xtw.WriteStartElement("protocolo"); // proto ini
			xtw.WriteAttributeString("nome", "Open field");
			this.CriaItemXml(ref xtw, "A", "Walking");
			this.CriaItemXml(ref xtw, "S", "Rearing");
			this.CriaItemXml(ref xtw, "D", "Grooming");
			this.CriaItemXml(ref xtw, "F", "Feaces");
			this.CriaItemXml(ref xtw, "H", string.Empty);
			this.CriaItemXml(ref xtw, "J", "Periphery");
			this.CriaItemXml(ref xtw, "K", "Centre");
			this.CriaItemXml(ref xtw, "L", string.Empty);
			xtw.WriteEndElement(); // proto fim

			xtw.WriteStartElement("protocolo"); // proto ini
			xtw.WriteAttributeString("nome", "Forced swim test");
			this.CriaItemXml(ref xtw, "A", "Climbing");
			this.CriaItemXml(ref xtw, "S", "Swimming");
			this.CriaItemXml(ref xtw, "D", "Headshake");
			this.CriaItemXml(ref xtw, "F", string.Empty);
			this.CriaItemXml(ref xtw, "H", string.Empty);
			this.CriaItemXml(ref xtw, "J", "Floating");
			this.CriaItemXml(ref xtw, "K", "Freezing");
			this.CriaItemXml(ref xtw, "L", "Dive");
			xtw.WriteEndElement(); // proto fim

			//xtw.WriteStartElement("protocolo"); // proto ini
			//xtw.WriteAttributeString("nome", "Reconhecimento de Objeto");
			//this.CriaItemXml(ref xtw, "A", "Inicio");
			//this.CriaItemXml(ref xtw, "S", "Obejto 1");
			//this.CriaItemXml(ref xtw, "D", "Obejto 2");
			//this.CriaItemXml(ref xtw, "F", "Obejto Novo");
			//this.CriaItemXml(ref xtw, "H", string.Empty);
			//this.CriaItemXml(ref xtw, "J", string.Empty);
			//this.CriaItemXml(ref xtw, "K", string.Empty);
			//this.CriaItemXml(ref xtw, "L", string.Empty);
			//xtw.WriteEndElement(); // proto fim

			//xtw.WriteStartElement("protocolo"); // proto ini
			//xtw.WriteAttributeString("nome", "Grooming Microestruturado");
			//this.CriaItemXml(ref xtw, "A", "Sem Grooming");
			//this.CriaItemXml(ref xtw, "S", "Lamber Patas");
			//this.CriaItemXml(ref xtw, "D", "Lavar Focinho");
			//this.CriaItemXml(ref xtw, "F", string.Empty);
			//this.CriaItemXml(ref xtw, "H", string.Empty);
			//this.CriaItemXml(ref xtw, "J", "Lavar Cabeça/Orelhas");
			//this.CriaItemXml(ref xtw, "K", "Coçar Corpo");
			//this.CriaItemXml(ref xtw, "L", "Lavar Patas Traseiras/Genitália");
			//xtw.WriteEndElement(); // proto fim
		}

		void CriaItemXml(ref XmlTextWriter xtw, string tecla, string nome)
		{
			xtw.WriteStartElement("item");
			xtw.WriteAttributeString("tecla", tecla);
			xtw.WriteString(nome);
			xtw.WriteEndElement();
		}

		void LimparComboBoxProtocolo()
		{
			// Limpar itens ComboBox
			if (cbTipoComportamento != null && cbTipoComportamento.Items.Count > 0)
			{
				int itensCombo = cbTipoComportamento.Items.Count;
				for (int i = itensCombo - 1; i >= 0; i--)
				{
					cbTipoComportamento.Items.RemoveAt(i);
				}
			}
		}

		enum VideoState
		{
			Played,
			Stopped,
			Paused
		}


	}

	public struct ComboProtocoloItem
	{
		public string Nome { get; set; }
		public XmlNodeList Nodos { get; set; }
	}

	public class TemposItem
	{
		public Key Tecla;
		public TimeSpan? Inicio;
		public TimeSpan? Fim;
		public int Ordem;
		public string TextoAtalho;

		public int Duracao()
		{
			if (Inicio.HasValue && Fim.HasValue)
				return (int)(Fim.Value - Inicio.Value).TotalSeconds;
			else
				return 0;
		}
	}
}
