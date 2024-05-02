using System.Collections.Generic;
using System.Linq;

namespace ExcelReport.Data
{
	internal class Report
	{
		/// <summary>
		/// Сотрудники со списком задач
		/// </summary>
		internal Dictionary<Employee, List<Data.Task>> EmployesTask;

		/// <summary>
		/// Отдел
		/// </summary>
		internal Department Department;

		/// <summary>
		/// Количество задач в отделе
		/// </summary>
		internal int TasksCount
		{
			get
			{
				if (EmployesTask == null)
				{
					return 0;
				}

				return EmployesTask.Values.SelectMany(tasks => tasks).Count();
			}
		}
	}
}
