using ExcelReport.MSOffice;
using System.Collections.Generic;

namespace ExcelReport.Interfaces
{
	interface ITableConvertor<T>
	{
		T ConvertTable(List<ExcelRow> list);
	}
}
