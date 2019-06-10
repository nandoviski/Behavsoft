using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Linq;
using System.Windows;
using Excel = Microsoft.Office.Interop.Excel;

namespace Behavsoft
{
	public class ExcelUtil
	{
		BackgroundWorker bwGerarExcel;
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
		}

		void bwGerarExcel_DoWork(object sender, DoWorkEventArgs e)
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
				xlWorkSheet.Cells[1, mesclarFrInicio] = "Frequency";
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
				xlWorkSheet.Cells[1, mesclarDrInicio] = "Duration";
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
				xlWorkSheet.Cells[1, 1] = "Key";
				((Excel.Range)xlWorkSheet.Cells[1, 1]).Font.Bold = true;
				xlWorkSheet.Cells[1, 2] = "Start";
				((Excel.Range)xlWorkSheet.Cells[1, 2]).Font.Bold = true;
				xlWorkSheet.Cells[1, 3] = "End";
				((Excel.Range)xlWorkSheet.Cells[1, 3]).Font.Bold = true;
				xlWorkSheet.Cells[1, 4] = "Duration (sec)";
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

				object fileFormat;
				if (caminhoCompleto.ToLower().EndsWith("xlsx"))
				{
					fileFormat = Excel.XlFileFormat.xlWorkbookDefault;
				}
				else
				{
					fileFormat = Excel.XlFileFormat.xlWorkbookNormal;
				}

				xlWorkBook.SaveAs(caminhoCompleto, fileFormat, misValue, misValue, misValue, misValue,
					Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
				xlWorkBook.Close(true, misValue, misValue);
				xlApp.Quit();

				liberarObjetos(xlWorkSheet);
				liberarObjetos(xlWorkBook);
				liberarObjetos(xlApp);

				//exibe mensagem ao usuario
				var result = MessageBox.Show("Excel generated successfully.\nDo you wish to open the file now?", "Success", MessageBoxButton.YesNo, MessageBoxImage.Information);
				if (result == MessageBoxResult.Yes)
				{
					Process.Start(caminhoCompleto);
				}

			}
			catch (Exception excpt)
			{
				MessageBox.Show(excpt.Message, "Error", MessageBoxButton.OK, MessageBoxImage.Error);
			}

		}

		void bwGerarExcel_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
		{
			if (JanelaPai != null)
				JanelaPai.MostrarUcCarregando(false);
		}

		void liberarObjetos(object obj)
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

		public void MergeExcel(string[] excelPath, Dictionary<string, string> nomeTecla, string savePath)
		{
			var listaDic = new Dictionary<string, List<KeyData>>();

			foreach (string item in excelPath)
			{
				var excel = new Excel.Application();
				var wb = excel.Workbooks.Open(item);

				var lista = new List<KeyData>();

				// Pega nome da planilha
				var planilha1 = string.Empty;
				foreach (Excel.Worksheet sh in wb.Worksheets)
				{
					planilha1 = (sh.Name);
					break;
				}

				var listaCelulas = "GHIJKLMNOPQRSTUVWXYZ";
				for (int i = 0; i < listaCelulas.Length; i++)
				{
					var celula = listaCelulas[i].ToString();
					var data = new KeyData
					{
						tecla = ((Excel.Range)((Excel.Worksheet)wb.Sheets[planilha1]).Cells[2, celula]).Value2
					};

					if (data.tecla == null)
						continue;

					var achou = BuscarDadosStruct(lista, data.tecla.ToString());

					if (achou < 0)
					{
						data.frequecia = ((Excel.Range)((Excel.Worksheet)wb.Sheets[planilha1]).Cells[3, celula]).Value2;
						lista.Add(data);
					}
					else
					{
						var keyfound = lista[achou];
						keyfound.duracao = ((Excel.Range)((Excel.Worksheet)wb.Sheets[planilha1]).Cells[3, celula]).Value2;
						lista.Insert(achou, keyfound);
						lista.RemoveAt(achou + 1);
					}
				}

				var indexInicio = item.LastIndexOf("\\");
				var indexFinal = item.LastIndexOf(".");

				var excelFile = item.Substring(indexInicio + 1, (indexFinal - indexInicio) - 1);
				listaDic.Add(excelFile, lista);

				wb.Close();
				excel.Quit();
			}

			GerarJuntaTabelas(listaDic, nomeTecla, savePath);
		}

		void GerarJuntaTabelas(Dictionary<string, List<KeyData>> dic, Dictionary<string, string> nomeTecla, string caminhoCompleto)
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
					foreach (KeyData item in dic[key])
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
				xlWorkSheet.Cells[1, 2] = "Frequency";
				((Excel.Range)xlWorkSheet.Cells[1, 2]).Font.Bold = true;
				((Excel.Range)xlWorkSheet.Cells[1, 2]).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;

				Excel.Range mesclarDu = xlWorkSheet.get_Range(xlWorkSheet.Cells[1, teclasUsadas.Count + 2], xlWorkSheet.Cells[1, teclasUsadas.Count + 1 + teclasUsadas.Count]);
				mesclarDu.Merge(true);
				xlWorkSheet.Cells[1, teclasUsadas.Count + 2] = "Duration";
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
					var teclasUsadasCount = teclasUsadas.Count;
					object frequecia, duracao;
					var dictionaryKeys = dic[key];

					foreach (var item in teclasUsadas)
					{
						if (dictionaryKeys.Any(k => k.tecla.ToString() == item))
						{
							var usedKey = dictionaryKeys.First(k => k.tecla.ToString() == item);
							frequecia = usedKey.frequecia;
							duracao = usedKey.duracao;
						}
						else
						{
							frequecia = 0;
							duracao = 0;
						}

						coluna = teclasUsadas.IndexOf(item) + 2;
						xlWorkSheet.Cells[linhaCont, coluna] = frequecia;
						((Excel.Range)xlWorkSheet.Cells[linhaCont, coluna]).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;

						xlWorkSheet.Cells[linhaCont, (coluna + teclasUsadasCount)] = duracao;
						((Excel.Range)xlWorkSheet.Cells[linhaCont, (coluna + teclasUsadasCount)]).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
					}

					linhaCont++;
				}

				object fileFormat;
				if (caminhoCompleto.ToLower().EndsWith("xlsx"))
				{
					fileFormat = Excel.XlFileFormat.xlWorkbookDefault;
				}
				else
				{
					fileFormat = Excel.XlFileFormat.xlWorkbookNormal;
				}

				xlWorkBook.SaveAs(caminhoCompleto, fileFormat, misValue, misValue, misValue, misValue,
					Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
				xlWorkBook.Close(true, misValue, misValue);
				xlApp.Quit();

				liberarObjetos(xlWorkSheet);
				liberarObjetos(xlWorkBook);
				liberarObjetos(xlApp);

				var result = MessageBox.Show("Merge completed successfully.\nDo you wish to open the file now?", "Success", MessageBoxButton.YesNo, MessageBoxImage.Information);
				if (result == MessageBoxResult.Yes)
				{
					Process.Start(caminhoCompleto);
				}
			}
			catch (Exception excpt)
			{
				MessageBox.Show(excpt.Message, "Error", MessageBoxButton.OK, MessageBoxImage.Error);
			}


		}

		int BuscarDadosStruct(List<KeyData> lista, string tecla)
		{
			int index = -1;

			for (int i = 0; i < lista.Count; i++)
			{
				if (lista[i].tecla.ToString() == tecla)
					return i;
			}

			return index;
		}

		struct KeyData
		{
			public object tecla;
			public object frequecia;
			public object duracao;
		}
	}
}
