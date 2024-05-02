using System.Collections.Generic;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelReport.MSOffice
{
	internal class ExcelReader
	{
		#region Properties
		/// <summary>
		/// Книга
		/// </summary>
		private Excel.Workbook _workBook;

		/// <summary>
		/// Листы
		/// </summary>
		internal Excel.Sheets _worksheets;

		/// <summary>
		/// Приложение
		/// </summary>
		internal Excel.Application _app;
		#endregion Properties

		#region Constructor
		internal ExcelReader()
		{

		}
		#endregion Constructor

		#region Methods
		/// <summary>
		/// Открыть Excel-файл
		/// </summary>
		/// <returns></returns>
		internal string OpenExcel()
		{
			// Выбрать путь и имя файла в диалоговом окне
			OpenFileDialog ofd = new OpenFileDialog();
			// Задаем расширение имени файла по умолчанию (открывается папка с программой)
			ofd.DefaultExt = "*.xls;*.xlsx;*.xlsb";
			// Задаем заголовок диалогового окна
			ofd.Title = "Выберите файл";

			if (ofd.ShowDialog() == DialogResult.OK) 
			{
				if (ofd.FileName.EndsWith(".xls") || 
					ofd.FileName.EndsWith(".xlsx") || 
					ofd.FileName.EndsWith(".xlsb"))
				{
					return ofd.FileName;
				}
				else
				{
					return string.Empty;
				}
			}
			return string.Empty;
		}

		/// <summary>
		/// Загрузить файл
		/// </summary>
		/// <param name="path"></param>
		internal void LoadExcel(string path)
		{
			_app = new Excel.Application();
			_workBook = _app.Workbooks.Open(path);
			_worksheets = _workBook.Sheets;
		}

		/// <summary>
		/// Чтение страницы
		/// </summary>
		/// <param name="nameSheet"></param>
		/// <returns></returns>
		internal List<ExcelRow> ReadExcelSheet(string nameSheet)
		{
			Excel.Worksheet worksheet = GetWorksheet(nameSheet);

			if (worksheet == null)
			{
				MessageBox.Show("Листа с именем " + nameSheet + "нет в данном файле", "Ошибка!");
				return null;
			}

			List<ExcelRow> table = new List<ExcelRow>();

			int rowCount = worksheet.UsedRange.Rows.Count;
			int сolumnCount = worksheet.UsedRange.Columns.Count;

			for (int rowNumber = 2; rowNumber <= rowCount; rowNumber++)
			{
				ExcelRow excelRow = new ExcelRow();

				for (int colNumber = 1; colNumber <= сolumnCount; colNumber++)
				{ 
					var cell = worksheet.Cells[rowNumber, colNumber];
					excelRow.Row.Add(cell.Text);
				}
				table.Add(excelRow);
			}

			return table;
		}

		/// <summary>
		/// получить лист по имени
		/// </summary>
		/// <param name="name"></param>
		/// <returns></returns>
		private Excel.Worksheet GetWorksheet(string name)
		{
			foreach (Excel.Worksheet worksheet in _worksheets)
			{
				if (worksheet.Name == name)
				{
					return worksheet;
				}
			}
			return null;
		}

		/// <summary>
		/// Закрыть файл
		/// </summary>
		internal void CloseExcel()
		{
			try
			{
				_workBook.Close();
				_app.Quit();

				System.Runtime.InteropServices.Marshal.ReleaseComObject(_workBook);
				System.Runtime.InteropServices.Marshal.ReleaseComObject(_app);
			}
			catch 
			{ 
				//информация для логирования
			}
		}
		#endregion Methods
	}
}
