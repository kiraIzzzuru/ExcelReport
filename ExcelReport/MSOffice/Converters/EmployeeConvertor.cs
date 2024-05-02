using ExcelReport.Data;
using ExcelReport.Interfaces;
using System;
using System.Collections.Generic;

namespace ExcelReport.MSOffice.Converters
{
	internal class EmployeeConvertor : ITableConvertor<List<Employee>>
	{
		public List<Employee> ConvertTable(List<ExcelRow> list)
		{
			List<Employee> employees = new List<Employee>();

			foreach (var row in list)
			{
				if (row.Row.Find(x => x.Length > 0) == null)
				{
					continue;
				}

				Employee employee = GetEmployeeFromExcelRow(row.Row);

				employees.Add(employee);
			}

			return employees;
		}

		private Employee GetEmployeeFromExcelRow(List<string> row)
		{
			Employee employee = new Employee();

			ulong.TryParse(row[0], out employee.Id);
			employee.SecondName = row[1];
			employee.Name = row[2];
			employee.SurName = row[3];

			DateTime.TryParse(row[4], out employee.BirthDay);
			uint.TryParse(row[5], out employee.DepartmentId);

			return employee;
		}
	}
}
