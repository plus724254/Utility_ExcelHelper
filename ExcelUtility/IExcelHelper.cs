using System;
using System.IO;

namespace ExcelUtilitys
{
    public interface IExcelHelper
    {
        void CreateSheet(string sheetName);
        int GetCellIndex();
        int GetRowIndex();
        int GetRowCount();
        void OnExport(Stream stream);
        void SetCellIndex(int cellIndex);
        void SetRowCellIndex(int rowIndex, int cellIndex);
        void SetRowIndex(int rowIndex);
        void SetSheet(string name);
        void SetSheetIndex(int index);
        void SetValue<T>(T value, bool nextCell = true);
        void NextRow(bool firstCell = true);
        void NextCell(bool firstRow = true);
        string GetCellValueString(bool nextCell = true);
        double GetCellValueNumber(bool nextCell = true);
        DateTime GetCellValueDateTime(bool nextCell = true);
    }
}