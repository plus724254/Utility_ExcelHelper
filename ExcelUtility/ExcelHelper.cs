using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
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
            return _cell.ColumnIndex;
        }

        public void SetRowCellIndex(int rowIndex, int cellIndex)
        {
            CheckOrCreateRowCell(rowIndex, cellIndex);
        }
        public void SetRowIndex(int rowIndex)
        {
            CheckOrCreateRowCell(rowIndex, _cell.ColumnIndex);
        }
        public void SetCellIndex(int cellIndex)
        {
            CheckOrCreateRowCell(_row.RowNum, cellIndex);
        }

        public void SetSheet(string name)
        {
            _sheet = _workbook.GetSheet(name);
            _row = _sheet.GetRow(0);
        }
        public void SetSheetIndex(int index)
        {
            _sheet = _workbook.GetSheetAt(index);
            _row = _sheet.GetRow(0);
        }

        public void SetValue<T>(T value, bool nextCell = false)
        {
            var valueType = typeof(T);
            // 配合C#6.0以下版本
            switch (Type.GetTypeCode(valueType))
            {
                case TypeCode.DateTime:
                    _cell.SetCellValue(Convert.ToDateTime(value));
                    break;
                case TypeCode.Int16:
                case TypeCode.Int32:
                case TypeCode.Int64:
                case TypeCode.Decimal:
                case TypeCode.Double:
                case TypeCode.Single:
                    _cell.SetCellValue(Convert.ToDouble(value));
                    break;
                case TypeCode.String:
                    _cell.SetCellValue(value.ToString());
                    break;
                default:
                    if (value != null)
                    {
                        if (valueType == typeof(DateTime?))
                        {
                            _cell.SetCellValue(Convert.ToDateTime(value));
                        }
                        else if (valueType == typeof(int?) || valueType == typeof(double?)
                            || valueType == typeof(decimal?) || valueType == typeof(long?)
                            || valueType == typeof(short?) || valueType == typeof(float?))
                        {
                            _cell.SetCellValue(Convert.ToDouble(value));
                        }
                        break;
                    }
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
