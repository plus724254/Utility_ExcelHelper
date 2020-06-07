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
            _cell = _row.GetCell(0);
        }

        public void CreateSheet(string sheetName)
        {
            _sheet = _workbook.CreateSheet(sheetName);
            CheckOrCreateRowCell(0, 0);
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

        public void SetValue<T>(T value, bool nextCell = true)
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

        public void NextRow(bool firstCell = true)
        {
            SetRowIndex(GetRowIndex() +1);
            if (firstCell)
            {
                SetCellIndex(0);
            }
        }

        public void NextCell(bool firstRow = true)
        {
            SetCellIndex(GetCellIndex() + 1);
            if (firstRow)
            {
                SetRowIndex(0);
            }
        }

        public string GetCellValueString(bool nextCell = true)
        {
            CheckOrCreateRowCell(_row.RowNum, _cell.ColumnIndex);
            var cellStringValue = _cell.StringCellValue;
            if (nextCell)
            {
                SetCellIndex(_cell.ColumnIndex + 1);
            }

            return cellStringValue;
        }

        public double GetCellValueNumber(bool nextCell = true)
        {
            CheckOrCreateRowCell(_row.RowNum, _cell.ColumnIndex);
            var cellNumericValue = _cell.NumericCellValue;
            if (nextCell)
            {
                SetCellIndex(_cell.ColumnIndex + 1);
            }

            return cellNumericValue;
        }

        public DateTime GetCellValueDateTime(bool nextCell = true)
        {
            CheckOrCreateRowCell(_row.RowNum, _cell.ColumnIndex);
            var cellDateTimeValue = _cell.DateCellValue;
            if (nextCell)
            {
                SetCellIndex(_cell.ColumnIndex + 1);
            }

            return cellDateTimeValue;
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
