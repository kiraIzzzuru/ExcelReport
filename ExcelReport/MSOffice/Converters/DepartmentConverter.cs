using ExcelReport.Data;
using ExcelReport.Interfaces;
using System.Collections.Generic;

namespace ExcelReport.MSOffice.Converters
{
	internal class DepartmentConverter : ITableConvertor<List<Department>>
	{
		public List<Department> ConvertTable(List<ExcelRow> list)
		{
			List<Department> departments = new List<Department>();

			foreach (var row in list)
			{
				if (row.Row.Find(x => x.Length > 0) == null)
				{
					continue;
				}

				Department department = GetDepartmentFromExcelRow(row.Row);

				departments.Add(department);
			}

			return departments;
		}

		private Department GetDepartmentFromExcelRow(List<string> row)
		{
			Department department = new Department();

			uint.TryParse(row[0], out department.Id);
			department.Name = row[1];

			return department;
		}
	}
}
