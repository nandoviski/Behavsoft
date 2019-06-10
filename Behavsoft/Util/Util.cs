using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Behavsoft
{
	public class Util
	{
		public static string FileDialogAllFilesFilter => "All files (*.*)|*.*";
		public static string FileDialogExcelFilesFilter => "Excel Workbook|*.xlsx|Excel 97-2003 Workbook|*.xls";
		public static string FileDialogExcelFilesFilterWithAllFiles => FileDialogExcelFilesFilter + "|" + FileDialogAllFilesFilter;
	}
}
