using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using UtilityLibrary.Loggers;

namespace EPPlus_Library
{
	public class EPPlusReader
	{
		private FileInfo existingFile;
		private ExcelPackage package;
		private ExcelWorksheet sheet;
		Logger logger = Logger.getInstance;
		public EPPlusReader(string file)
		{
			try
			{
				existingFile = new FileInfo(file);
				package = new ExcelPackage(existingFile);
			}
			catch (Exception ex)
			{
				logger.logException(ex);
			}
		}
		public void SetSheet(string sheetName)
		{
			sheet = package.Workbook.Worksheets[sheetName];
		}
		public void SetSheet(int sheetNumber)
		{
			sheet = package.Workbook.Worksheets[sheetNumber];
		}
		public object GetCell(int row, int column)
		{
			try
			{
				return sheet.Cells[row, column].Value;
			}
			catch (Exception ex)
			{
				logger.logException(ex);
			}
			return "Error trying to get Cell at the specified row: " + row + " column: " + column;
		}
		public Tuple<int,int> GetFirstPopulatedRow(string firstHeader)
		{
			for(int row = 1; row < 30; row++)
			{
				for (int column = 1; column < 30; column++)
				{
					object cellValue = sheet.Cells[row, column].Value;
					if (cellValue != null && cellValue.ToString().Trim() == firstHeader)
					{
						Tuple<int,int> firstCell =  new Tuple<int, int>( row, column );
						return firstCell;
					}
				}
			}
			return new Tuple<int, int>(-1, -1);
		}
		public List<List<object>> GetTable()
		{
			List<List<object>> table = new List<List<object>>();
			try
			{
				int lastColumn = sheet.Dimension.End.Column;
				int lastRow = sheet.Dimension.End.Row;
				for (int row = 1; row <= lastRow; row++)
				{
					List<object> rowList = new List<object>();
					for (int column = 1; column <= lastColumn; column++)
					{
						object cell = sheet.Cells[row, column].Value;
						rowList.Add(cell);
					}
					table.Add(rowList);
				}
			}
			catch(Exception ex)
			{
				logger.logException(ex);
			}
			return table;
		}
		public List<List<object>> GetTable(int firstRow, int firstColumn, int OffSetRowMinus)
		{
			List<List<object>> table = new List<List<object>>();
			try
			{
				int lastColumn = sheet.Dimension.End.Column;
				int lastRow = sheet.Dimension.End.Row;
				for (int row = firstRow; row <= lastRow- OffSetRowMinus; row++)
				{
					List<object> rowList = new List<object>();
					for (int column = firstColumn; column <= lastColumn; column++)
					{
						object cell = sheet.Cells[row, column].Value;
						rowList.Add(cell);
					}
					table.Add(rowList);
				}
			}
			catch (Exception ex)
			{
				logger.logException(ex);
			}
			return table;
		}
		public void ReleaseMemory()
		{
			try
			{
				sheet.Dispose();
				package.Dispose();
			}
			catch(Exception ex)
			{
				logger.logException(ex);
			}
		}
	}
}
