using System;
using System.Collections.Generic;
using System.Text;
using System.IO;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.Windows;
using System.Data;
using Excel = Microsoft.Office.Interop.Excel;
using System.ComponentModel;

namespace Behavsoft
{
    public class ExcelUtil
    {
        //public static void Exportar(string caminhoCompleto, List<TemposItem> tempos)
        //{
        //    // criar um arquivo para escrever
        //    using (StreamWriter sw = File.CreateText(caminhoCompleto))
        //    {
        //        try
        //        {
        //            List<ItemCalculo> frequencia = new List<ItemCalculo>();
        //            int index;
        //            foreach (TemposItem item in tempos)
        //            {
        //                index = -1;
        //                for (int i = 0; i < frequencia.Count; i++)
        //                {
        //                    if (frequencia[i].Tecla == item.Tecla)
        //                    {
        //                        index = i;
        //                        break;
        //                    }
        //                }

        //                if (index >= 0)
        //                {
        //                    ItemCalculo novo = frequencia[index];
        //                    novo.Frequencia++;
        //                    novo.TotalTempo += item.Duracao();
        //                    frequencia[index] = novo;
        //                }
        //                else
        //                {
        //                    frequencia.Add(new ItemCalculo() { Frequencia = 1, TotalTempo = item.Duracao(), Tecla = item.Tecla });
        //                }
        //            }
        //            //sw.WriteLine("Tecla" + "\t" + "Frequencia" + "\t" + "Duracao total");
        //            //foreach (var item in temp)
        //            //{
        //            //    sw.WriteLine(item.Tecla.ToString() + "\t" + item.Frequencia + "\t" + item.TotalTempo);
        //            //}

        //            sw.WriteLine("Tecla" + "\t" + "Inicio" + "\t" + "Fim" + "\t" + "Duracao (seg)" + "" + "\t" + "" + "\t" + "Tecla" + "\t" + "Frequencia" + "\t" + "Duracao Total (seg)");

        //            // percorre o datareader e escreve os dados no arquivo .xls definido
        //            for (int i = 0; i < tempos.Count; i++)
        //            {
        //                if (i >= frequencia.Count)
        //                    sw.WriteLine(tempos[i].Tecla.ToString() + "\t" + tempos[i].Inicio + "\t" + tempos[i].Fim + "\t" + tempos[i].Duracao());
        //                else
        //                    sw.WriteLine(tempos[i].Tecla.ToString() + "\t" + tempos[i].Inicio + "\t" + tempos[i].Fim + "\t" + tempos[i].Duracao() + " " + "\t" + " " + "\t" + frequencia[i].Tecla + "\t" + frequencia[i].Frequencia + "\t" + frequencia[i].TotalTempo);
        //            }

        //            //exibe mensagem ao usuario
        //            MessageBox.Show("Arquivo " + caminhoCompleto + " gerado com sucesso.");
        //        }
        //        catch (Exception excpt)
        //        {
        //            MessageBox.Show(excpt.Message);
        //        }
        //    }
        //}

        private BackgroundWorker bwGerarExcel;
        public MainWindow JanelaPai = null; 

        public void GerarExcel(string caminhoCompleto, List<TemposItem> tempos)
        {
            object[] param = new object[2];
            param[0] = caminhoCompleto;
            param[1] = tempos;

            if (JanelaPai != null)
                JanelaPai.MostrarUcCarregando();

            bwGerarExcel = new BackgroundWorker();
            bwGerarExcel.DoWork += bwGerarExcel_DoWork;
            bwGerarExcel.RunWorkerCompleted += bwGerarExcel_RunWorkerCompleted;
            bwGerarExcel.RunWorkerAsync(param);

            //try
            //{
            //    Excel.Application xlApp;
            //    Excel.Workbook xlWorkBook;
            //    Excel.Worksheet xlWorkSheet;
            //    object misValue = System.Reflection.Missing.Value;

            //    xlApp = new Excel.Application();
            //    xlWorkBook = xlApp.Workbooks.Add(misValue);

            //    xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);


            //    List<ItemCalculo> frequencia = new List<ItemCalculo>();
            //    int index;
            //    foreach (TemposItem item in tempos)
            //    {
            //        index = -1;
            //        for (int i = 0; i < frequencia.Count; i++)
            //        {
            //            if (frequencia[i].Tecla == item.Tecla)
            //            {
            //                index = i;
            //                break;
            //            }
            //        }

            //        if (index >= 0)
            //        {
            //            ItemCalculo novo = frequencia[index];
            //            novo.Frequencia++;
            //            novo.TotalTempo += item.Duracao();
            //            frequencia[index] = novo;
            //        }
            //        else
            //        {
            //            frequencia.Add(new ItemCalculo() { Frequencia = 1, TotalTempo = item.Duracao(), Tecla = item.Tecla, TextoAtalho = item.TextoAtalho });
            //        }
            //    }

            //    // Monta frequencia
            //    int colunaFr = 7;
            //    int mesclarFrInicio, mesclarFrFim;
            //    mesclarFrInicio = colunaFr;
            //    mesclarFrFim = colunaFr;
            //    foreach (var item in frequencia)
            //    {
            //        xlWorkSheet.Cells[2, colunaFr] = item.TextoAtalho;
            //        ((Excel.Range)xlWorkSheet.Cells[2, colunaFr]).Font.Bold = true;
            //        ((Excel.Range)xlWorkSheet.Cells[2, colunaFr]).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
            //        //((Excel.Range)xlWorkSheet.Cells[2, colunaFr]).Borders[Excel.XlBordersIndex.xlEdgeBottom].Color = 0x000000;
            //        //((Excel.Range)xlWorkSheet.Cells[2, colunaFr]).Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;

            //        xlWorkSheet.Cells[3, colunaFr] = item.Frequencia;
            //        ((Excel.Range)xlWorkSheet.Cells[3, colunaFr]).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
            //        ((Excel.Range)xlWorkSheet.Cells[3, colunaFr]).Borders[Excel.XlBordersIndex.xlEdgeBottom].Color = 0x000000;
            //        ((Excel.Range)xlWorkSheet.Cells[3, colunaFr]).Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
            //        mesclarFrFim = colunaFr;
            //        colunaFr++;
            //    }
            //    //Mesclar
            //    Excel.Range mesclarFr = xlWorkSheet.get_Range(xlWorkSheet.Cells[1, mesclarFrInicio], xlWorkSheet.Cells[1, mesclarFrFim]);
            //    mesclarFr.Merge(true);
            //    xlWorkSheet.Cells[1, mesclarFrInicio] = "Frequência";
            //    ((Excel.Range)xlWorkSheet.Cells[1, mesclarFrInicio]).Font.Bold = true;
            //    ((Excel.Range)xlWorkSheet.Cells[1, mesclarFrInicio]).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
            //    //Borda
            //    ((Excel.Range)xlWorkSheet.Cells[1, mesclarFrInicio]).Borders[Excel.XlBordersIndex.xlEdgeLeft].Color = 0x000000;
            //    ((Excel.Range)xlWorkSheet.Cells[1, mesclarFrInicio]).Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
            //    ((Excel.Range)xlWorkSheet.Cells[2, mesclarFrInicio]).Borders[Excel.XlBordersIndex.xlEdgeLeft].Color = 0x000000;
            //    ((Excel.Range)xlWorkSheet.Cells[2, mesclarFrInicio]).Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
            //    ((Excel.Range)xlWorkSheet.Cells[3, mesclarFrInicio]).Borders[Excel.XlBordersIndex.xlEdgeLeft].Color = 0x000000;
            //    ((Excel.Range)xlWorkSheet.Cells[3, mesclarFrInicio]).Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
            //    ((Excel.Range)xlWorkSheet.Cells[1, mesclarFrFim]).Borders[Excel.XlBordersIndex.xlEdgeRight].Color = 0x000000;
            //    ((Excel.Range)xlWorkSheet.Cells[1, mesclarFrFim]).Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
            //    ((Excel.Range)xlWorkSheet.Cells[2, mesclarFrFim]).Borders[Excel.XlBordersIndex.xlEdgeRight].Color = 0x000000;
            //    ((Excel.Range)xlWorkSheet.Cells[2, mesclarFrFim]).Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
            //    ((Excel.Range)xlWorkSheet.Cells[3, mesclarFrFim]).Borders[Excel.XlBordersIndex.xlEdgeRight].Color = 0x000000;
            //    ((Excel.Range)xlWorkSheet.Cells[3, mesclarFrFim]).Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;

            //    // Monta duração
            //    int mesclarDrInicio, mesclarDrFim;
            //    mesclarDrInicio = colunaFr;
            //    mesclarDrFim = colunaFr;
            //    foreach (var item in frequencia)
            //    {
            //        xlWorkSheet.Cells[2, colunaFr] = item.TextoAtalho;
            //        ((Excel.Range)xlWorkSheet.Cells[2, colunaFr]).Font.Bold = true;
            //        ((Excel.Range)xlWorkSheet.Cells[2, colunaFr]).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                    
            //        xlWorkSheet.Cells[3, colunaFr] = item.TotalTempo;
            //        ((Excel.Range)xlWorkSheet.Cells[3, colunaFr]).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
            //        ((Excel.Range)xlWorkSheet.Cells[3, colunaFr]).Borders[Excel.XlBordersIndex.xlEdgeBottom].Color = 0x000000;
            //        ((Excel.Range)xlWorkSheet.Cells[3, colunaFr]).Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
            //        mesclarDrFim = colunaFr;
            //        colunaFr++;
            //    }
            //    Excel.Range mesclarDu = xlWorkSheet.get_Range(xlWorkSheet.Cells[1, mesclarDrInicio], xlWorkSheet.Cells[1, mesclarDrFim]);
            //    mesclarDu.Merge(true);
            //    xlWorkSheet.Cells[1, mesclarDrInicio] = "Duração";
            //    ((Excel.Range)xlWorkSheet.Cells[1, mesclarDrInicio]).Font.Bold = true;
            //    ((Excel.Range)xlWorkSheet.Cells[1, mesclarDrInicio]).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
            //    //Borda
            //    ((Excel.Range)xlWorkSheet.Cells[1, mesclarDrInicio]).Borders[Excel.XlBordersIndex.xlEdgeLeft].Color = 0x000000;
            //    ((Excel.Range)xlWorkSheet.Cells[1, mesclarDrInicio]).Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
            //    ((Excel.Range)xlWorkSheet.Cells[2, mesclarDrInicio]).Borders[Excel.XlBordersIndex.xlEdgeLeft].Color = 0x000000;
            //    ((Excel.Range)xlWorkSheet.Cells[2, mesclarDrInicio]).Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
            //    ((Excel.Range)xlWorkSheet.Cells[3, mesclarDrInicio]).Borders[Excel.XlBordersIndex.xlEdgeLeft].Color = 0x000000;
            //    ((Excel.Range)xlWorkSheet.Cells[3, mesclarDrInicio]).Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
            //    ((Excel.Range)xlWorkSheet.Cells[1, mesclarDrFim]).Borders[Excel.XlBordersIndex.xlEdgeRight].Color = 0x000000;
            //    ((Excel.Range)xlWorkSheet.Cells[1, mesclarDrFim]).Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
            //    ((Excel.Range)xlWorkSheet.Cells[2, mesclarDrFim]).Borders[Excel.XlBordersIndex.xlEdgeRight].Color = 0x000000;
            //    ((Excel.Range)xlWorkSheet.Cells[2, mesclarDrFim]).Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
            //    ((Excel.Range)xlWorkSheet.Cells[3, mesclarDrFim]).Borders[Excel.XlBordersIndex.xlEdgeRight].Color = 0x000000;
            //    ((Excel.Range)xlWorkSheet.Cells[3, mesclarDrFim]).Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;

            //    // Monta contagem teclas
            //    int linhaCont = 2;
            //    xlWorkSheet.Cells[1, 1] = "Tecla";
            //    ((Excel.Range)xlWorkSheet.Cells[1, 1]).Font.Bold = true;
            //    xlWorkSheet.Cells[1, 2] = "Início";
            //    ((Excel.Range)xlWorkSheet.Cells[1, 2]).Font.Bold = true;
            //    xlWorkSheet.Cells[1, 3] = "Fim";
            //    ((Excel.Range)xlWorkSheet.Cells[1, 3]).Font.Bold = true;
            //    xlWorkSheet.Cells[1, 4] = "Duração (seg)";
            //    ((Excel.Range)xlWorkSheet.Cells[1, 4]).Font.Bold = true;
            //    ((Excel.Range)xlWorkSheet.Columns[4]).AutoFit();
                
            //    for (int i = 0; i < tempos.Count; i++)
            //    {
            //        xlWorkSheet.Cells[linhaCont, 1] = tempos[i].TextoAtalho;
            //        xlWorkSheet.Cells[linhaCont, 2] = tempos[i].Inicio.Value.Hours.ToString("00") + ":" + tempos[i].Inicio.Value.Minutes.ToString("00") + ":" + tempos[i].Inicio.Value.Seconds.ToString("00");
            //        xlWorkSheet.Cells[linhaCont, 3] = tempos[i].Fim.Value.Hours.ToString("00") + ":" + tempos[i].Fim.Value.Minutes.ToString("00") + ":" + tempos[i].Fim.Value.Seconds.ToString("00");
            //        xlWorkSheet.Cells[linhaCont, 4] = tempos[i].Duracao();
            //        linhaCont++;
            //    }

            //    xlWorkBook.SaveAs(caminhoCompleto, Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue,
            //        Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
            //    xlWorkBook.Close(true, misValue, misValue);
            //    xlApp.Quit();

            //    liberarObjetos(xlWorkSheet);
            //    liberarObjetos(xlWorkBook);
            //    liberarObjetos(xlApp);

            //    //exibe mensagem ao usuario
            //    MessageBox.Show("Arquivo " + caminhoCompleto + " gerado com sucesso.", "Sucesso", MessageBoxButton.OK, MessageBoxImage.Information);
            //}
            //catch (Exception excpt)
            //{
            //    MessageBox.Show(excpt.Message, "Erro", MessageBoxButton.OK, MessageBoxImage.Error);
            //}


        }

        private void bwGerarExcel_DoWork(object sender, DoWorkEventArgs e)
        {
            object[] param = e.Argument as object[];

            string caminhoCompleto = param[0] as string;
            List<TemposItem> tempos = param[1] as List<TemposItem>;

            try
            {
                Excel.Application xlApp;
                Excel.Workbook xlWorkBook;
                Excel.Worksheet xlWorkSheet;
                object misValue = System.Reflection.Missing.Value;

                xlApp = new Excel.Application();
                xlWorkBook = xlApp.Workbooks.Add(misValue);

                xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);


                List<ItemCalculo> frequencia = new List<ItemCalculo>();
                int index;
                foreach (TemposItem item in tempos)
                {
                    index = -1;
                    for (int i = 0; i < frequencia.Count; i++)
                    {
                        if (frequencia[i].Tecla == item.Tecla)
                        {
                            index = i;
                            break;
                        }
                    }

                    if (index >= 0)
                    {
                        ItemCalculo novo = frequencia[index];
                        novo.Frequencia++;
                        novo.TotalTempo += item.Duracao();
                        frequencia[index] = novo;
                    }
                    else
                    {
                        frequencia.Add(new ItemCalculo() { Frequencia = 1, TotalTempo = item.Duracao(), Tecla = item.Tecla, TextoAtalho = item.TextoAtalho });
                    }
                }

                // Monta frequencia
                int colunaFr = 7;
                int mesclarFrInicio, mesclarFrFim;
                mesclarFrInicio = colunaFr;
                mesclarFrFim = colunaFr;
                foreach (var item in frequencia)
                {
                    xlWorkSheet.Cells[2, colunaFr] = item.TextoAtalho;
                    ((Excel.Range)xlWorkSheet.Cells[2, colunaFr]).Font.Bold = true;
                    ((Excel.Range)xlWorkSheet.Cells[2, colunaFr]).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                    //((Excel.Range)xlWorkSheet.Cells[2, colunaFr]).Borders[Excel.XlBordersIndex.xlEdgeBottom].Color = 0x000000;
                    //((Excel.Range)xlWorkSheet.Cells[2, colunaFr]).Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;

                    xlWorkSheet.Cells[3, colunaFr] = item.Frequencia;
                    ((Excel.Range)xlWorkSheet.Cells[3, colunaFr]).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                    ((Excel.Range)xlWorkSheet.Cells[3, colunaFr]).Borders[Excel.XlBordersIndex.xlEdgeBottom].Color = 0x000000;
                    ((Excel.Range)xlWorkSheet.Cells[3, colunaFr]).Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                    mesclarFrFim = colunaFr;
                    colunaFr++;
                }
                //Mesclar
                Excel.Range mesclarFr = xlWorkSheet.get_Range(xlWorkSheet.Cells[1, mesclarFrInicio], xlWorkSheet.Cells[1, mesclarFrFim]);
                mesclarFr.Merge(true);
                xlWorkSheet.Cells[1, mesclarFrInicio] = "Frequência";
                ((Excel.Range)xlWorkSheet.Cells[1, mesclarFrInicio]).Font.Bold = true;
                ((Excel.Range)xlWorkSheet.Cells[1, mesclarFrInicio]).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                //Borda
                ((Excel.Range)xlWorkSheet.Cells[1, mesclarFrInicio]).Borders[Excel.XlBordersIndex.xlEdgeLeft].Color = 0x000000;
                ((Excel.Range)xlWorkSheet.Cells[1, mesclarFrInicio]).Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                ((Excel.Range)xlWorkSheet.Cells[2, mesclarFrInicio]).Borders[Excel.XlBordersIndex.xlEdgeLeft].Color = 0x000000;
                ((Excel.Range)xlWorkSheet.Cells[2, mesclarFrInicio]).Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                ((Excel.Range)xlWorkSheet.Cells[3, mesclarFrInicio]).Borders[Excel.XlBordersIndex.xlEdgeLeft].Color = 0x000000;
                ((Excel.Range)xlWorkSheet.Cells[3, mesclarFrInicio]).Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                ((Excel.Range)xlWorkSheet.Cells[1, mesclarFrFim]).Borders[Excel.XlBordersIndex.xlEdgeRight].Color = 0x000000;
                ((Excel.Range)xlWorkSheet.Cells[1, mesclarFrFim]).Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                ((Excel.Range)xlWorkSheet.Cells[2, mesclarFrFim]).Borders[Excel.XlBordersIndex.xlEdgeRight].Color = 0x000000;
                ((Excel.Range)xlWorkSheet.Cells[2, mesclarFrFim]).Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                ((Excel.Range)xlWorkSheet.Cells[3, mesclarFrFim]).Borders[Excel.XlBordersIndex.xlEdgeRight].Color = 0x000000;
                ((Excel.Range)xlWorkSheet.Cells[3, mesclarFrFim]).Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;

                // Monta duração
                int mesclarDrInicio, mesclarDrFim;
                mesclarDrInicio = colunaFr;
                mesclarDrFim = colunaFr;
                foreach (var item in frequencia)
                {
                    xlWorkSheet.Cells[2, colunaFr] = item.TextoAtalho;
                    ((Excel.Range)xlWorkSheet.Cells[2, colunaFr]).Font.Bold = true;
                    ((Excel.Range)xlWorkSheet.Cells[2, colunaFr]).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;

                    xlWorkSheet.Cells[3, colunaFr] = item.TotalTempo;
                    ((Excel.Range)xlWorkSheet.Cells[3, colunaFr]).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                    ((Excel.Range)xlWorkSheet.Cells[3, colunaFr]).Borders[Excel.XlBordersIndex.xlEdgeBottom].Color = 0x000000;
                    ((Excel.Range)xlWorkSheet.Cells[3, colunaFr]).Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                    mesclarDrFim = colunaFr;
                    colunaFr++;
                }
                Excel.Range mesclarDu = xlWorkSheet.get_Range(xlWorkSheet.Cells[1, mesclarDrInicio], xlWorkSheet.Cells[1, mesclarDrFim]);
                mesclarDu.Merge(true);
                xlWorkSheet.Cells[1, mesclarDrInicio] = "Duração";
                ((Excel.Range)xlWorkSheet.Cells[1, mesclarDrInicio]).Font.Bold = true;
                ((Excel.Range)xlWorkSheet.Cells[1, mesclarDrInicio]).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                //Borda
                ((Excel.Range)xlWorkSheet.Cells[1, mesclarDrInicio]).Borders[Excel.XlBordersIndex.xlEdgeLeft].Color = 0x000000;
                ((Excel.Range)xlWorkSheet.Cells[1, mesclarDrInicio]).Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                ((Excel.Range)xlWorkSheet.Cells[2, mesclarDrInicio]).Borders[Excel.XlBordersIndex.xlEdgeLeft].Color = 0x000000;
                ((Excel.Range)xlWorkSheet.Cells[2, mesclarDrInicio]).Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                ((Excel.Range)xlWorkSheet.Cells[3, mesclarDrInicio]).Borders[Excel.XlBordersIndex.xlEdgeLeft].Color = 0x000000;
                ((Excel.Range)xlWorkSheet.Cells[3, mesclarDrInicio]).Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                ((Excel.Range)xlWorkSheet.Cells[1, mesclarDrFim]).Borders[Excel.XlBordersIndex.xlEdgeRight].Color = 0x000000;
                ((Excel.Range)xlWorkSheet.Cells[1, mesclarDrFim]).Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                ((Excel.Range)xlWorkSheet.Cells[2, mesclarDrFim]).Borders[Excel.XlBordersIndex.xlEdgeRight].Color = 0x000000;
                ((Excel.Range)xlWorkSheet.Cells[2, mesclarDrFim]).Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                ((Excel.Range)xlWorkSheet.Cells[3, mesclarDrFim]).Borders[Excel.XlBordersIndex.xlEdgeRight].Color = 0x000000;
                ((Excel.Range)xlWorkSheet.Cells[3, mesclarDrFim]).Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;

                // Monta contagem teclas
                int linhaCont = 2;
                xlWorkSheet.Cells[1, 1] = "Tecla";
                ((Excel.Range)xlWorkSheet.Cells[1, 1]).Font.Bold = true;
                xlWorkSheet.Cells[1, 2] = "Início";
                ((Excel.Range)xlWorkSheet.Cells[1, 2]).Font.Bold = true;
                xlWorkSheet.Cells[1, 3] = "Fim";
                ((Excel.Range)xlWorkSheet.Cells[1, 3]).Font.Bold = true;
                xlWorkSheet.Cells[1, 4] = "Duração (seg)";
                ((Excel.Range)xlWorkSheet.Cells[1, 4]).Font.Bold = true;
                ((Excel.Range)xlWorkSheet.Columns[4]).AutoFit();

                for (int i = 0; i < tempos.Count; i++)
                {
                    xlWorkSheet.Cells[linhaCont, 1] = tempos[i].TextoAtalho;
                    xlWorkSheet.Cells[linhaCont, 2] = tempos[i].Inicio.Value.Hours.ToString("00") + ":" + tempos[i].Inicio.Value.Minutes.ToString("00") + ":" + tempos[i].Inicio.Value.Seconds.ToString("00");
                    xlWorkSheet.Cells[linhaCont, 3] = tempos[i].Fim.Value.Hours.ToString("00") + ":" + tempos[i].Fim.Value.Minutes.ToString("00") + ":" + tempos[i].Fim.Value.Seconds.ToString("00");
                    xlWorkSheet.Cells[linhaCont, 4] = tempos[i].Duracao();
                    linhaCont++;
                }

                xlWorkBook.SaveAs(caminhoCompleto, Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue,
                    Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
                xlWorkBook.Close(true, misValue, misValue);
                xlApp.Quit();

                liberarObjetos(xlWorkSheet);
                liberarObjetos(xlWorkBook);
                liberarObjetos(xlApp);

                //exibe mensagem ao usuario
                MessageBox.Show("Arquivo " + caminhoCompleto + " gerado com sucesso.", "Sucesso", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception excpt)
            {
                MessageBox.Show(excpt.Message, "Erro", MessageBoxButton.OK, MessageBoxImage.Error);
            }

        }

        void bwGerarExcel_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if (JanelaPai != null)
                JanelaPai.MostrarUcCarregando(false);
        }

        private void liberarObjetos(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
                MessageBox.Show("Ocorreu um erro durante a liberação do objeto " + ex.ToString());
            }
            finally
            {
                GC.Collect();
            }
        }

        struct ItemCalculo
        {
            public System.Windows.Input.Key Tecla;
            public string TextoAtalho;
            public int TotalTempo;
            public int Frequencia;
        }

        public void MesclarExcel(string[] excelPath, Dictionary<string, string> nomeTecla, string savePath, string tipoExcel)
        {
            Dictionary<string, List<tempStruct>> listaDic = new Dictionary<string, List<tempStruct>>();

            foreach (string item in excelPath)
            {
                Excel.Application excel = new Excel.Application();
                Excel.Workbook wb = excel.Workbooks.Open(item);

                List<tempStruct> lista = new List<tempStruct>();

                // Pega nome da planilha
                string planilha1 = string.Empty;
                foreach (Excel.Worksheet sh in wb.Worksheets)
                {
                    planilha1 = (sh.Name);
                    break;
                }

                if (tipoExcel.ToUpper() == "NOVO")
                {
                    // PLANILHAS MODELO NOVO
                    // Sempre linha 3 e inicia na celula 'G'
                    string listaCelulas = "GHIJKLMNOPQRSTUVWXYZ";
                    for (int i = 0; i < listaCelulas.Length; i++)
                    {
                        string celula = listaCelulas[i].ToString();
                        tempStruct aa = new tempStruct();
                        aa.tecla = ((Excel.Range)((Excel.Worksheet)wb.Sheets[planilha1]).Cells[2, celula]).Value2;

                        if (aa.tecla == null)
                            break;

                        int acho = BuscarDadosStruct(lista, aa.tecla.ToString());

                        if (acho < 0)
                        {
                            aa.frequecia = ((Excel.Range)((Excel.Worksheet)wb.Sheets[planilha1]).Cells[3, celula]).Value2;
                            lista.Add(aa);
                        }
                        else
                        {
                            tempStruct a = lista[acho];
                            a.duracao = ((Excel.Range)((Excel.Worksheet)wb.Sheets[planilha1]).Cells[3, celula]).Value2;
                            lista.Insert(acho, a);
                            lista.RemoveAt(acho + 1);
                        }
                    }
                }
                else if (tipoExcel.ToUpper() == "ANTIGO")
                {
                    // PLANILHAS MODELO ANTIGOS
                    // Começa no 2 pois é onde tem os valores
                    for (int i = 2; i < 100; i++)
                    {
                        tempStruct aa = new tempStruct();
                        aa.tecla = ((Excel.Range)((Excel.Worksheet)wb.Sheets[planilha1]).Cells[i, "F"]).Value2;

                        if (aa.tecla != null)
                        {
                            aa.frequecia = ((Excel.Range)((Excel.Worksheet)wb.Sheets[planilha1]).Cells[i, "G"]).Value2;
                            aa.duracao = ((Excel.Range)((Excel.Worksheet)wb.Sheets[planilha1]).Cells[i, "H"]).Value2;
                        }
                        else
                        {
                            aa.tecla = ((Excel.Range)((Excel.Worksheet)wb.Sheets[planilha1]).Cells[i, "G"]).Value2;
                            aa.frequecia = ((Excel.Range)((Excel.Worksheet)wb.Sheets[planilha1]).Cells[i, "H"]).Value2;
                            aa.duracao = ((Excel.Range)((Excel.Worksheet)wb.Sheets[planilha1]).Cells[i, "I"]).Value2;
                        }

                        if (aa.tecla == null && aa.frequecia == null && aa.duracao == null)
                            break;

                        lista.Add(aa);
                    }
                }

                int indexInicio = item.LastIndexOf("\\");
                int indexFinal = item.LastIndexOf(".");

                string aaaa = item.Substring(indexInicio + 1, (indexFinal - indexInicio) - 1);
                listaDic.Add(aaaa, lista);

                wb.Close();
                excel.Quit();
            }

            GerarJuntaTabelas(listaDic, nomeTecla, savePath);
        }

        private void GerarJuntaTabelas(Dictionary<string, List<tempStruct>> dic, Dictionary<string, string> nomeTecla, string caminhoCompleto)
        {
            try
            {
                Excel.Application xlApp;
                Excel.Workbook xlWorkBook;
                Excel.Worksheet xlWorkSheet;
                object misValue = System.Reflection.Missing.Value;

                xlApp = new Excel.Application();
                xlWorkBook = xlApp.Workbooks.Add(misValue);

                xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

                // Acha todas teclas (OBS: Order teclas aqui)
                List<string> teclasUsadas = new List<string>();
                foreach (string key in dic.Keys)
                {
                    foreach (tempStruct item in dic[key])
                    {
                        if (!teclasUsadas.Contains(item.tecla.ToString()))
                            teclasUsadas.Add(item.tecla.ToString());
                    }
                }

                // Seta ordem das teclas
                List<string> teclasUsadasOrdenada = new List<string>();
                foreach (var item in nomeTecla.Values)
                {
                    if (teclasUsadas.Contains(item))
                        teclasUsadasOrdenada.Add(item);
                }
                teclasUsadas = teclasUsadasOrdenada;

                // Monta contagem teclas
                int linhaCont = 3;

                Excel.Range mesclarFr = xlWorkSheet.get_Range(xlWorkSheet.Cells[1, 2], xlWorkSheet.Cells[1, teclasUsadas.Count + 1]);
                mesclarFr.Merge(true);
                xlWorkSheet.Cells[1, 2] = "Frequência";
                ((Excel.Range)xlWorkSheet.Cells[1, 2]).Font.Bold = true;
                ((Excel.Range)xlWorkSheet.Cells[1, 2]).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;

                Excel.Range mesclarDu = xlWorkSheet.get_Range(xlWorkSheet.Cells[1, teclasUsadas.Count + 2], xlWorkSheet.Cells[1, teclasUsadas.Count + 1 + teclasUsadas.Count]);
                mesclarDu.Merge(true);
                xlWorkSheet.Cells[1, teclasUsadas.Count + 2] = "Duração";
                ((Excel.Range)xlWorkSheet.Cells[1, teclasUsadas.Count + 2]).Font.Bold = true;
                ((Excel.Range)xlWorkSheet.Cells[1, teclasUsadas.Count + 2]).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                
                int colinaTeclas = 2;
                foreach (var item in teclasUsadas)
                {
                    xlWorkSheet.Cells[2, colinaTeclas] = item;
                    ((Excel.Range)xlWorkSheet.Columns[colinaTeclas]).AutoFit();
                    ((Excel.Range)xlWorkSheet.Cells[2, colinaTeclas]).Font.Bold = true;
                    xlWorkSheet.Cells[2, colinaTeclas + teclasUsadas.Count] = item;
                    ((Excel.Range)xlWorkSheet.Columns[colinaTeclas + teclasUsadas.Count]).AutoFit();
                    ((Excel.Range)xlWorkSheet.Cells[2, colinaTeclas + teclasUsadas.Count]).Font.Bold = true;
                    colinaTeclas++;
                }
                
                foreach (string key in dic.Keys)
                {
                    xlWorkSheet.Cells[linhaCont, 1] = key;

                    int coluna;
                    foreach (tempStruct item in dic[key])
                    {
                        coluna = teclasUsadas.IndexOf(item.tecla.ToString()) + 2;
                        xlWorkSheet.Cells[linhaCont, coluna] = item.frequecia;
                        ((Excel.Range)xlWorkSheet.Cells[linhaCont, coluna]).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                    }
                    foreach (tempStruct item in dic[key])
                    {
                        coluna = teclasUsadas.IndexOf(item.tecla.ToString()) + 2 + teclasUsadas.Count;
                        xlWorkSheet.Cells[linhaCont, coluna] = item.duracao;
                        ((Excel.Range)xlWorkSheet.Cells[linhaCont, coluna]).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                    }

                    linhaCont++;
                }

                xlWorkBook.SaveAs(caminhoCompleto, Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue,
                    Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
                xlWorkBook.Close(true, misValue, misValue);
                xlApp.Quit();

                liberarObjetos(xlWorkSheet);
                liberarObjetos(xlWorkBook);
                liberarObjetos(xlApp);

                //exibe mensagem ao usuario
                MessageBox.Show("Arquivo " + caminhoCompleto + " gerado com sucesso.", "Sucesso", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception excpt)
            {
                MessageBox.Show(excpt.Message, "Erro", MessageBoxButton.OK, MessageBoxImage.Error);
            }


        }

        private int BuscarDadosStruct(List<tempStruct> lista, string tecla)
        {
            int index = -1;

            for (int i = 0; i < lista.Count; i++)
            {
                if (lista[i].tecla.ToString() == tecla)
                    return i;
            }

            return index;
        }

        private struct tempStruct
        {
            public object tecla;
            public object frequecia;
            public object duracao;
        }
    }
}
