using ExcelReport.Data;
using Microsoft.Office.Interop.Word;
using System.Collections.Generic;
using Word = Microsoft.Office.Interop.Word;

namespace ExcelReport.MSOffice
{
	internal class WordWriter
	{
		#region Methods
		internal static void CreateWordDocument(List<Report> report)
		{
			Word.Application wordApp = new Word.Application();
			Word.Document doc = wordApp.Documents.Add();

			string caption = "Отчет по загрузке";
			Paragraph paraAboveTable = doc.Content.Paragraphs.Add();
			paraAboveTable.Range.Text = caption;
			wordApp.Selection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;

			paraAboveTable.Range.InsertParagraphAfter();

			// Добавление таблицы
			Word.Table table = doc.Tables.Add(paraAboveTable.Range, 1, 2);

			// Заполнение ячеек таблицы
			table.Cell(1, 1).Range.Text = "Отдел";
			table.Cell(1, 2).Range.Text = "Количество задач";

			FillTable(table, report);

			table.Borders.Enable = 1; // Включаем границы
			table.Borders.InsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;
			table.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;

			// Отображение документа
			doc.Activate();
			wordApp.Visible = true;
			wordApp.Activate();
		}

		/// <summary>
		/// Наполнить таблицу данными
		/// </summary>
		private static void FillTable(Word.Table table, List<Report> report)
		{
			string currentDepartment = "";

			foreach (var data in report)
			{
				if (!currentDepartment.Equals(data.Department.Name))
				{
					Word.Row row = table.Rows.Add();
					row.Cells[1].Range.Text = data.Department.Name;
					row.Cells[2].Range.Text = data.TasksCount.ToString();

					foreach (Word.Cell cell in row.Cells)
					{
						cell.Shading.BackgroundPatternColor = Word.WdColor.wdColorGray25;
					}
				}

				foreach (var employe in data.EmployesTask)
				{
					Word.Row newRow = table.Rows.Add();
					newRow.Cells[1].Range.Text = employe.Key.EmployeName;
					newRow.Cells[2].Range.Text = employe.Value.Count.ToString();

					foreach (Word.Cell cell in newRow.Cells)
					{
						cell.Shading.BackgroundPatternColor = Word.WdColor.wdColorWhite;
					}
				}
			}
		}
		#endregion Methods
	}
}
