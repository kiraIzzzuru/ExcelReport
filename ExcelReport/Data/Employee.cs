using System;

namespace ExcelReport.Data
{
	internal class Employee
	{
		/// <summary>
		/// Идентификатор
		/// </summary>
		internal ulong Id;

		/// <summary>
		/// Фамилия
		/// </summary>
		internal string SecondName;

		/// <summary>
		/// Имя
		/// </summary>
		internal string Name;

		/// <summary>
		/// Отчество
		/// </summary>
		internal string SurName;

		/// <summary>
		/// Дата рождения
		/// </summary>
		internal DateTime BirthDay;

		/// <summary>
		/// Идентификаторо отдела
		/// </summary>
		internal uint DepartmentId;

		/// <summary>
		/// Имя отрудника в формате ФИО
		/// </summary>
		internal string EmployeName
		{
			get 
			{
				string value = SecondName + " " + Name?.Substring(0, 1) + ".";

				if (!string.IsNullOrEmpty(SurName) && SurName.Length > 0)
				{
					value += SurName?.Substring(0, 1) + ".";
				}

				return value;
			}
		}
	}
}
