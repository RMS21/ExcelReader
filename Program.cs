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

		public IEnumerable<DataRow> getData(){
			return this.getRows ("Sheet1");
		}


	}

	class Program{
		public static void Main(String [] args){
			
			ExcelData e = new ExcelData (//path to file);
			e.getData ();
		}
	}
}