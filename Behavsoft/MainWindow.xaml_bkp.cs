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
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Windows.Threading;

namespace Behavsoft
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private TimeSpan? ultimoClique;
        private Key ultimaTecla;
        private List<TemposItem> tempos;
        private VideoState estadoVideo;
        private DispatcherTimer timer;
        private bool isDragging;
        private TimeSpan? inicioAnalise;

        public MainWindow()
        {
            InitializeComponent();

            KeyDown += MainWindow_KeyDown;
            this.WindowState = System.Windows.WindowState.Maximized;

            tempos = new List<TemposItem>();
            estadoVideo = VideoState.Stopped;

            cbTipoComportamento.SelectionChanged += cbTipoComportamento_SelectionChanged;
            cbTipoComportamento.Items.Add("Labirinto Em Cruz Elevado");
            cbTipoComportamento.Items.Add("Campo Aberto");
            cbTipoComportamento.Items.Add("Nado Forçado");

            cbTipoComportamento.SelectedIndex = 0;
            inicioAnalise = null;

            isDragging = false;
            timer = new DispatcherTimer();
            timer.Interval = TimeSpan.FromMilliseconds(300);
            timer.Tick += timer_Tick;
        }

        private void timer_Tick(object sender, EventArgs e)
        {
            if(!isDragging)
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

        private void MainWindow_KeyDown(object sender, KeyEventArgs e)
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
                    TimeSpan tempoAgora = new TimeSpan(VideoControl.Position.Hours, VideoControl.Position.Minutes, VideoControl.Position.Seconds);

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
                        TemposItem a = BuscarUltimoTempo(e.Key);
                        if (a == null)
                        {
                            tempos.Add(new TemposItem() { Tecla = e.Key, Inicio = tempoAgora, TextoAtalho = BuscarTextoAtalho(e.Key) });
                            this.MudarCorTecla(e.Key, true);
                        }
                        else
                        {
                            a.Fim = tempoAgora;
                            this.MudarCorTecla(e.Key, false);
                        }

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
                JuntarExcel j = new JuntarExcel();
                j.ShowDialog();
            }
        }

        private TemposItem BuscarUltimoTempo(Key tecla)
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

        private void BrowseButton_Click_1(object sender, RoutedEventArgs e)
        {
            try
            {
                Microsoft.Win32.OpenFileDialog openDlg = new Microsoft.Win32.OpenFileDialog();
                openDlg.InitialDirectory = @"c:\";
                bool? ret = openDlg.ShowDialog();

                if (ret.HasValue && ret.Value)
                {
                    MediaPathTextBox.Text = openDlg.FileName;
                    VideoControl.Source = new Uri(openDlg.FileName);
                    estadoVideo = VideoState.Stopped;
                    tempos = new List<TemposItem>();


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
                MessageBox.Show("Erro ao carregar o vídeo", "Erro", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void PlayButton_Click_1(object sender, RoutedEventArgs e)
        {
            if (MediaPathTextBox.Text.Length <= 0)
            {
                MessageBox.Show("Você deve selecionar um video", "Atenção", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            if (estadoVideo == VideoState.Stopped)
            {
                txtTempoAnalise.Foreground = new SolidColorBrush(Colors.Black);
                txtTempoAnalise.FontWeight = FontWeights.Normal;
                txtTempoAnalise.Text = "00:00:00";
                tempos = new List<TemposItem>();
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

        private void PauseButton_Click_1(object sender, RoutedEventArgs e)
        {
            VideoControl.Pause();
            estadoVideo = VideoState.Paused;
            timer.Stop();
        }

        private void StopButton_Click_1(object sender, RoutedEventArgs e)
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

        private void MostrarTempos()
        {
            lbAcao.Items.Clear();

            foreach (var item in tempos)
            {
                string msg = item.Tecla + " " + item.Inicio.Value;//.ToString(@"hh\:mm\:ss");

                if (item.Fim.HasValue)
                {
                    msg += " " + item.Fim.Value + " " + item.Duracao(); //.ToString(@"hh\:mm\:ss")
                }

                lbAcao.Items.Add(msg);
            }

            if(lbAcao.Items.Count > 0)
                lbAcao.ScrollIntoView(lbAcao.Items[lbAcao.Items.Count - 1]);
        }

        private void BloquearComponentesConfig(bool bloquear)
        {
            BrowseButton.IsEnabled = !bloquear;
            gbTeclas.IsEnabled = !bloquear;
            cbTipoComportamento.IsEnabled = !bloquear;
        }

        private void btnSalvarComo_Click(object sender, RoutedEventArgs e)
        {
            if (estadoVideo != VideoState.Stopped)
            {
                MessageBox.Show("Não é possivel salvar enquanto o video estiver rodando", "Atenção", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            if (tempos == null || tempos.Count < 1)
            {
                MessageBox.Show("Não existe dados a serem salvos", "Atenção", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            Microsoft.Win32.SaveFileDialog sabeDlg = new Microsoft.Win32.SaveFileDialog();
            sabeDlg.InitialDirectory = @"c:\";
            sabeDlg.Filter = "Pasta de Trabalho do Excel 97-2003|*.xls|Pasta de Trabalho do Excel|*.xlsx";
            bool? ret = sabeDlg.ShowDialog();

            if (ret.HasValue && ret.Value)
            {
                ExcelUtil util = new ExcelUtil();
                util.JanelaPai = this;
                util.GerarExcel(sabeDlg.FileName, tempos);
            }
        }

        private void sldBarraTempo_DragStarted_1(object sender, System.Windows.Controls.Primitives.DragStartedEventArgs e)
        {
            isDragging = true;
        }

        private void sldBarraTempo_ValueChanged(object sender, RoutedPropertyChangedEventArgs<double> e)
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

        private void sldBarraTempo_DragCompleted_1(object sender, System.Windows.Controls.Primitives.DragCompletedEventArgs e)
        {
            isDragging = false;
            //VideoControl.Position = TimeSpan.FromSeconds(sldBarraTempo.Value);
        }

        // Finaliza todas as teclas aberta a o clicar no Stop
        private void FinalizarTodasTeclas(TimeSpan final)
        {
            TimeSpan tempo = new TimeSpan(final.Hours, final.Minutes, final.Seconds);

            foreach (TemposItem item in tempos)
            {
                this.MudarCorTecla(item.Tecla, false);
                if (!item.Fim.HasValue)
                    item.Fim = tempo;
            }
        }

        private string BuscarTextoAtalho(Key tecla)
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

        private void MudarCorTecla(Key tecla, bool isPressionado)
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
                        tbTeclaA.Style = (Style)FindResource("TeclaAtalhoPressionada");
                    else
                        tbTeclaA.Style = (Style)FindResource("TeclaAtalhoNormal");
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

        enum VideoState
        {
            Played,
            Stopped,
            Paused
        }

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
