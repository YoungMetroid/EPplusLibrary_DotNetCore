using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPlus_Library
{
    public class EPPlusCreator
    {
        private ExcelPackage package;
        private ExcelWorksheet sheet;

        public EPPlusCreator()
        {
            package = new ExcelPackage();

        }
        public void CreateSheet(string sheetName)
        {
            package.Workbook.Worksheets.Add(sheetName);
        }
        public void SetSheet(string sheetName)
        {
            sheet = package.Workbook.Worksheets[sheetName];
        }
        public void SetSheet(int sheetNumber)
        {
            sheet = package.Workbook.Worksheets[sheetNumber];
        }

        public void SaveFile(string filePath, bool createAsMacro)
        {
            if(!createAsMacro)
                package.SaveAs(new FileInfo(filePath));
            else 
            {
                package.Workbook.CreateVBAProject();
                package.SaveAs(new FileInfo(filePath));
            }
        }
        public void WriteInfo(List<List<object>> table)
        {
            for (int row = 0; row < table.Count; row++)
            {
                for (int column = 0; column < table[row].Count; column++)
                {
                    sheet.Cells[row+1, column+1].Value = table[row][column];
                }
            }
        }
        public void WriteInfo(List<List<object>> table, int rowAvailableCell, int startingCol, int startingReadColumn, int lastReadColumn)
        {
            for (int row = 0; row < table.Count; row++)
            {
                for (int column = 0, dataStartColumn = startingReadColumn; dataStartColumn <= lastReadColumn; dataStartColumn++, column++)
                {
                    sheet.Cells[rowAvailableCell + row + 1, startingCol+ column + 1].Value = table[row][dataStartColumn];
                }
            }
        }
    }
}
