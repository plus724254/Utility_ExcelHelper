using System.IO;

namespace ExcelUtilitys
{
    public interface IExcelHelper
    {
        void CreateSheet(string sheetName);
        int GetCellIndex();
        int GetRowIndex();
        void OnExport(Stream stream);
        void SetCell(int cellIndex);
        void SetRow(int rowIndex);
        void SetRowCell(int rowIndex, int cellIndex);
        void SetSheet(string name);
        void SetSheetAt(int index);
        void SetValue<T>(T value, bool nextCell = false);
    }
}