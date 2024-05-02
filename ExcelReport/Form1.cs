using ExcelReport.Data;
using ExcelReport.Interfaces;
using ExcelReport.MSOffice;
using ExcelReport.MSOffice.Converters;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ExcelReport
{
	public partial class FrmReport : Form
	{
		public FrmReport()
		{
			InitializeComponent();
		}

		#region Methods
		/// <summary>
		/// Получить сконвертированные данные
		/// </summary>
		/// <typeparam name="T"></typeparam>
		/// <param name="tableConvertor"></param>
		/// <param name="list"></param>
		/// <returns></returns>
		private T GetTableDataFormat<T>(ITableConvertor<T> tableConvertor, List<ExcelRow> list)
		{
			return tableConvertor.ConvertTable(list);
		}

		/// <summary>
		/// Собрать отчет
		/// </summary>
		/// <returns></returns>
		private List<Report> CreateReport(List<Employee> employes, List<Department> departments, List<Data.Task> tasks)
		{
			List<Report> reports = new List<Report>();

			foreach (var department in departments)
			{
				Report report = new Report();

				report.Department = department;

				report.EmployesTask = employes.Where(e => e.DepartmentId == department.Id)
					.OrderByDescending(e => tasks.Where(x => x.Number == e.Id).ToList().Count)
					.ToDictionary(e => e, e => tasks.Where(x => x.Number == e.Id).ToList());

				reports.Add(report);
			}

			return reports.OrderByDescending(x => x.TasksCount).ToList();
		}

		/// <summary>
		/// Загрузка данных из Excel
		/// </summary>
		private void LoadExcelData(ExcelReader excelReader, string path)
		{
			excelReader.LoadExcel(path);

			var dataSheet1 = excelReader.ReadExcelSheet("Сотрудники");
			var dataSheet2 = excelReader.ReadExcelSheet("Отделы");
			var dataSheet3 = excelReader.ReadExcelSheet("Задачи");

			var employes = GetTableDataFormat<List<Employee>>(new EmployeeConvertor(), dataSheet1);
			var departments = GetTableDataFormat<List<Department>>(new DepartmentConverter(), dataSheet2);
			var tasks = GetTableDataFormat<List<Data.Task>>(new TaskConverter(), dataSheet3);

			WordWriter.CreateWordDocument(CreateReport(employes, departments, tasks));

			excelReader.CloseExcel();
		}

		#endregion Methods

		#region PrivateEvents
		private async void btnLoadExcel_Click(object sender, EventArgs e)
		{
			btnLoadExcel.Enabled = false;
			btnLoadExcel.Text = "Идет обработка файла";

			ExcelReader excelReader = new ExcelReader();

			string path = excelReader.OpenExcel();

			if (!string.IsNullOrEmpty(path))
			{
				await System.Threading.Tasks.Task.Factory.StartNew(() => LoadExcelData(excelReader, path), TaskCreationOptions.LongRunning)
					.ContinueWith(t =>
					{ }, TaskScheduler.FromCurrentSynchronizationContext());
			}

			btnLoadExcel.Enabled = true;
			btnLoadExcel.Text = "Запустить";
		}
		#endregion PrivateEvents
	}
}
