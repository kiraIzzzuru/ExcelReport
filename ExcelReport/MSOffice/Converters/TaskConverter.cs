using ExcelReport.Data;
using ExcelReport.Interfaces;
using System.Collections.Generic;

namespace ExcelReport.MSOffice.Converters
{
	internal class TaskConverter : ITableConvertor<List<Task>>
	{
		public List<Task> ConvertTable(List<ExcelRow> list)
		{
			List<Task> tasks = new List<Task>();

			foreach (var row in list)
			{
				if (row.Row.Find(x => x.Length > 0) == null)
				{
					continue;
				}

				Task task = GetTaskFromExcelRow(row.Row);

				tasks.Add(task);
			}

			return tasks;
		}

		private Task GetTaskFromExcelRow(List<string> row)
		{
			Task task = new Task();

			ulong.TryParse(row[0], out task.Id);
			ulong.TryParse(row[1], out task.Number);

			return task;
		}
	}
}
