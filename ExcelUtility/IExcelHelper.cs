using System.IO;

namespace ExcelUtilitys
{
    public interface IExcelHelper
    {
        void CreateSheet(string sheetName);
        int GetCellIndex();
        int GetRowIndex();
        void OnExport(Stream stream);
        void SetCellIndex(int cellIndex);
        void SetRowCellIndex(int rowIndex, int cellIndex);
        void SetRowIndex(int rowIndex);
        void SetSheet(string name);
        void SetSheetIndex(int index);
        void SetValue<T>(T value, bool nextCell = true);
    }
}