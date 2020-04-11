using System;
using System.IO;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace ExcelUtilitys
{
    public class ExcelHelper : IExcelHelper
    {
        private IWorkbook _workbook;
        private ISheet _sheet;
        private IRow _row;
        private ICell _cell;

        public ExcelHelper()
        {
            this._workbook = new XSSFWorkbook();
        }
        public ExcelHelper(Stream stream)
        {
            _workbook = new XSSFWorkbook(stream);
            _sheet = _workbook.GetSheetAt(0);
            _row = _sheet.GetRow(0);
        }

        public void CreateSheet(string sheetName)
        {
            this._sheet = this._workbook.CreateSheet(sheetName);
        }
        public int GetRowIndex()
        {
            return _row.RowNum;
        }
        public int GetCellIndex()
        {
            return _row.RowNum;
        }

        public void SetRowCell(int rowIndex, int cellIndex)
        {
            CheckOrCreateRowCell(rowIndex, cellIndex);
        }
        public void SetRow(int rowIndex)
        {
            CheckOrCreateRowCell(rowIndex, _cell.ColumnIndex);
        }
        public void SetCell(int cellIndex)
        {
            CheckOrCreateRowCell(_row.RowNum, cellIndex);
        }

        public void SetSheet(string name)
        {
            _sheet = _workbook.GetSheet(name);
            _row = _sheet.GetRow(0);
        }
        public void SetSheetAt(int index)
        {
            _sheet = _workbook.GetSheetAt(index);
            _row = _sheet.GetRow(0);
        }

        public void SetValue<T>(T value, bool nextCell = false)
        {
            switch (Type.GetTypeCode(typeof(T)))
            {
                case TypeCode.DateTime:
                    _cell.SetCellValue((DateTime)(object)value);
                    break;
                case TypeCode.UInt16:
                case TypeCode.UInt32:
                case TypeCode.UInt64:
                case TypeCode.Int16:
                case TypeCode.Int32:
                case TypeCode.Int64:
                case TypeCode.Decimal:
                case TypeCode.Double:
                case TypeCode.Single:
                    _cell.SetCellValue(Convert.ToDouble(value));
                    break;
                default:
                    _cell.SetCellValue(value.ToString());
                    break;
            }

            if (nextCell == true)
            {
                CheckOrCreateRowCell(_row.RowNum, _cell.ColumnIndex + 1);
            }
        }

        private void CheckOrCreateRowCell(int rowIndex, int cellIndex)
        {
            _row = _sheet.GetRow(rowIndex);
            if (_row == null)
            {
                _row = _sheet.CreateRow(rowIndex);
            }
            if (_row.GetCell(cellIndex) == null)
            {
                _row.CreateCell(cellIndex);
            }
            _cell = _row.GetCell(cellIndex);

            return;
        }

        public void OnExport(Stream stream)
        {
            _workbook.Write(stream);
            return;
        }
    }
}
