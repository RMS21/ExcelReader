﻿using Excel;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;

namespace ExcelDataReader
{
	public class ExcelData
	{
		private string _path;
		private List<List<string>> _rows = new List<List<string>>();


		public ExcelData(string path)
		{
			_path = path;
		} 

		private IExcelDataReader getExcelReader()
		{
			FileStream stream = File.Open(_path, FileMode.Open, FileAccess.Read);

			IExcelDataReader reader = null;
			try
			{
				if(_path.EndsWith(".xls"))
				{
					reader = ExcelReaderFactory.CreateBinaryReader(stream);
				}
				if(_path.EndsWith(".xlsx"))
				{
					reader = ExcelReaderFactory.CreateOpenXmlReader(stream);
				}
				return reader;
			}
			catch (Exception)
			{
				throw;
			}
		}

		private IEnumerable<DataRow> getRows(string sheet, bool firstRowIsColumnNames = true)
		{
			var reader = this.getExcelReader();
			reader.IsFirstRowAsColumnNames = firstRowIsColumnNames;
			var workSheet = reader.AsDataSet().Tables[sheet];
			var rows = from DataRow a in workSheet.Rows select a;
			return rows;
		}

		public List<List<string>> getData(){

			var rows = this.getRows("Sheet1");

			foreach (var row in rows)
			{
				List<string> list = new List<string>();

				for (int i = 0; i <= row.ItemArray.GetUpperBound(0); i++)
				{
					list.Add(row[i].ToString());
				}
				_rows.Add (list);
			}
			return _rows;
		}


	}

	class Program{
		public static void Main(String [] args){
			
			ExcelData e = new ExcelData ("/home/rasoul/a.xlsx");
			List<List<string>> temp = e.getData ();
			foreach(List<string> row in temp){
				foreach (string item in row) {
					Console.WriteLine (item);
				}	
			}
		}
	}
}