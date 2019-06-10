using System;
using System.Collections.Generic;
using System.IO;
using System.Windows;
using OfficeOpenXml;

namespace Behavsoft
{
	public class EPPlusUtil
	{
		public void GerarExcelTESTE(string caminhoCompleto, List<TemposItem> tempos)
		{
			//    object[] param = e.Argument as object[];

			//    string caminhoCompleto = param[0] as string;
			//    List<TemposItem> tempos = param[1] as List<TemposItem>;

			try
			{
				//FileInfo newFile = new FileInfo(caminhoCompleto);

				using (ExcelPackage pck = new ExcelPackage())
				{
					//Add the Content sheet
					var xlWorkSheet = pck.Workbook.Worksheets.Add("Plan1");

					//Excel.Application xlApp;
					//Excel.Workbook xlWorkBook;
					//Excel.Worksheet xlWorkSheet;
					//object misValue = System.Reflection.Missing.Value;

					//xlApp = new Excel.Application();
					//xlWorkBook = xlApp.Workbooks.Add(misValue);

					//xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);


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
						xlWorkSheet.Cells[2, colunaFr].Value = item.TextoAtalho;
						((ExcelRange)xlWorkSheet.Cells[2, colunaFr]).Style.Font.Bold = true;
						((ExcelRange)xlWorkSheet.Cells[2, colunaFr]).Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;

						xlWorkSheet.Cells[3, colunaFr].Value = item.Frequencia;
						((ExcelRange)xlWorkSheet.Cells[3, colunaFr]).Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
						//((ExcelRange)xlWorkSheet.Cells[3, colunaFr]).Style.Border[Excel.XlBordersIndex.xlEdgeBottom].Color = 0x000000;
						//((ExcelRange)xlWorkSheet.Cells[3, colunaFr]).Style.Border[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
						mesclarFrFim = colunaFr;
						colunaFr++;
					}
					//Mesclar
					xlWorkSheet.Cells[1, mesclarFrInicio, 1, mesclarFrFim].Merge = true;
					xlWorkSheet.Cells[1, mesclarFrInicio].Value = "Frequency";
					((ExcelRange)xlWorkSheet.Cells[1, mesclarFrInicio]).Style.Font.Bold = true;
					((ExcelRange)xlWorkSheet.Cells[1, mesclarFrInicio]).Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
					//Borda
					//((ExcelRange)xlWorkSheet.Cells[1, mesclarFrInicio]).Style.Border[Excel.XlBordersIndex.xlEdgeLeft].Color = 0x000000;
					//((ExcelRange)xlWorkSheet.Cells[1, mesclarFrInicio]).Style.Border[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
					//((ExcelRange)xlWorkSheet.Cells[2, mesclarFrInicio]).Style.Border[Excel.XlBordersIndex.xlEdgeLeft].Color = 0x000000;
					//((ExcelRange)xlWorkSheet.Cells[2, mesclarFrInicio]).Style.Border[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
					//((ExcelRange)xlWorkSheet.Cells[3, mesclarFrInicio]).Style.Border[Excel.XlBordersIndex.xlEdgeLeft].Color = 0x000000;
					//((ExcelRange)xlWorkSheet.Cells[3, mesclarFrInicio]).Style.Border[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
					//((ExcelRange)xlWorkSheet.Cells[1, mesclarFrFim]).Style.Border[Excel.XlBordersIndex.xlEdgeRight].Color = 0x000000;
					//((ExcelRange)xlWorkSheet.Cells[1, mesclarFrFim]).Style.Border[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
					//((ExcelRange)xlWorkSheet.Cells[2, mesclarFrFim]).Style.Border[Excel.XlBordersIndex.xlEdgeRight].Color = 0x000000;
					//((ExcelRange)xlWorkSheet.Cells[2, mesclarFrFim]).Style.Border[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
					//((ExcelRange)xlWorkSheet.Cells[3, mesclarFrFim]).Style.Border[Excel.XlBordersIndex.xlEdgeRight].Color = 0x000000;
					//((ExcelRange)xlWorkSheet.Cells[3, mesclarFrFim]).Style.Border[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;

					// Monta duração
					int mesclarDrInicio, mesclarDrFim;
					mesclarDrInicio = colunaFr;
					mesclarDrFim = colunaFr;
					foreach (var item in frequencia)
					{
						xlWorkSheet.Cells[2, colunaFr].Value = item.TextoAtalho;
						((ExcelRange)xlWorkSheet.Cells[2, colunaFr]).Style.Font.Bold = true;
						((ExcelRange)xlWorkSheet.Cells[2, colunaFr]).Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;

						xlWorkSheet.Cells[3, colunaFr].Value = item.TotalTempo;
						((ExcelRange)xlWorkSheet.Cells[3, colunaFr]).Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
						//((ExcelRange)xlWorkSheet.Cells[3, colunaFr]).Style.Border[Excel.XlBordersIndex.xlEdgeBottom].Color = 0x000000;
						//((ExcelRange)xlWorkSheet.Cells[3, colunaFr]).Style.Border[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
						mesclarDrFim = colunaFr;
						colunaFr++;
					}
					xlWorkSheet.Cells[1, mesclarDrInicio, 1, mesclarDrFim].Merge = true;
					xlWorkSheet.Cells[1, mesclarDrInicio].Value = "Duration";
					((ExcelRange)xlWorkSheet.Cells[1, mesclarDrInicio]).Style.Font.Bold = true;
					((ExcelRange)xlWorkSheet.Cells[1, mesclarDrInicio]).Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
					//Borda
					//((ExcelRange)xlWorkSheet.Cells[1, mesclarDrInicio]).Style.Border[Excel.XlBordersIndex.xlEdgeLeft].Color = 0x000000;
					//((ExcelRange)xlWorkSheet.Cells[1, mesclarDrInicio]).Style.Border[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
					//((ExcelRange)xlWorkSheet.Cells[2, mesclarDrInicio]).Style.Border[Excel.XlBordersIndex.xlEdgeLeft].Color = 0x000000;
					//((ExcelRange)xlWorkSheet.Cells[2, mesclarDrInicio]).Style.Border[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
					//((ExcelRange)xlWorkSheet.Cells[3, mesclarDrInicio]).Style.Border[Excel.XlBordersIndex.xlEdgeLeft].Color = 0x000000;
					//((ExcelRange)xlWorkSheet.Cells[3, mesclarDrInicio]).Style.Border[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
					//((ExcelRange)xlWorkSheet.Cells[1, mesclarDrFim]).Style.Border[Excel.XlBordersIndex.xlEdgeRight].Color = 0x000000;
					//((ExcelRange)xlWorkSheet.Cells[1, mesclarDrFim]).Style.Border[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
					//((ExcelRange)xlWorkSheet.Cells[2, mesclarDrFim]).Style.Border[Excel.XlBordersIndex.xlEdgeRight].Color = 0x000000;
					//((ExcelRange)xlWorkSheet.Cells[2, mesclarDrFim]).Style.Border[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
					//((ExcelRange)xlWorkSheet.Cells[3, mesclarDrFim]).Style.Border[Excel.XlBordersIndex.xlEdgeRight].Color = 0x000000;
					//((ExcelRange)xlWorkSheet.Cells[3, mesclarDrFim]).Style.Border[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;

					// Monta contagem teclas
					int linhaCont = 2;
					xlWorkSheet.Cells[1, 1].Value = "Key";
					((ExcelRange)xlWorkSheet.Cells[1, 1]).Style.Font.Bold = true;
					xlWorkSheet.Cells[1, 2].Value = "Start";
					((ExcelRange)xlWorkSheet.Cells[1, 2]).Style.Font.Bold = true;
					xlWorkSheet.Cells[1, 3].Value = "End";
					((ExcelRange)xlWorkSheet.Cells[1, 3]).Style.Font.Bold = true;
					xlWorkSheet.Cells[1, 4].Value = "Duration (sec)";
					((ExcelRange)xlWorkSheet.Cells[1, 4]).Style.Font.Bold = true;
					xlWorkSheet.Column(4).AutoFit();

					for (int i = 0; i < tempos.Count; i++)
					{
						xlWorkSheet.Cells[linhaCont, 1].Value = tempos[i].TextoAtalho;
						xlWorkSheet.Cells[linhaCont, 2].Value = tempos[i].Inicio.Value.Hours.ToString("00") + ":" + tempos[i].Inicio.Value.Minutes.ToString("00") + ":" + tempos[i].Inicio.Value.Seconds.ToString("00");
						xlWorkSheet.Cells[linhaCont, 3].Value = tempos[i].Fim.Value.Hours.ToString("00") + ":" + tempos[i].Fim.Value.Minutes.ToString("00") + ":" + tempos[i].Fim.Value.Seconds.ToString("00");
						xlWorkSheet.Cells[linhaCont, 4].Value = tempos[i].Duracao();
						linhaCont++;
					}

					Byte[] bin = pck.GetAsByteArray();
					File.WriteAllBytes(caminhoCompleto, bin);
				}

				//exibe mensagem ao usuario
				MessageBox.Show("Arquivo " + caminhoCompleto + " gerado com sucesso.", "Sucesso", MessageBoxButton.OK, MessageBoxImage.Information);
			}
			catch (Exception excpt)
			{
				MessageBox.Show(excpt.Message, "Erro", MessageBoxButton.OK, MessageBoxImage.Error);
			}

		}


		struct ItemCalculo
		{
			public System.Windows.Input.Key Tecla;
			public string TextoAtalho;
			public int TotalTempo;
			public int Frequencia;
		}

	}
}
