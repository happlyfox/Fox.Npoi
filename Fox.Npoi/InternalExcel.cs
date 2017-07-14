using Fox.Npoi.Core;
using Fox.Npoi.Ext;
using Fox.Npoi.Style;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace Fox.Npoi
{
    public interface IInternalExcel : IWrite
    {
        IWorkbook GetWorkbook();
    }

    public class InternalExcel : IInternalExcel
    {
        private IWorkbook _workbook;

        public InternalExcel(IExcelStyle excelStyle = null, string filePath = null)
        {
            if (string.IsNullOrWhiteSpace(filePath))
            {
                _workbook = new XSSFWorkbook();
            }
            else
            {
                using (var fs = new FileStream(filePath, FileMode.Open, FileAccess.Read))
                {
                    _workbook = WorkbookFactory.Create(fs);
                }
            }
            if (excelStyle != null)
                ExcelStyleUtil.RegisterStyle(excelStyle, _workbook);
        }

        public IWorkbook GetWorkbook()
        {
            return _workbook;
        }

        public void CreateSheet(string name)
        {
            _workbook.CreateSheet(name);
        }

        public void CreateRow(int sheetIndex, int rowIndex)
        {
            _workbook.GetSheetAt(sheetIndex).CreateRow(rowIndex);
        }

        public void WriteValue(int sheetIndex, int rowIndex, int cellIndex, dynamic value, int styleIndex, string formula = null)
        {
            var currentCell = _workbook.GetSheetAt(sheetIndex).GetRow(rowIndex).CreateCell(cellIndex);

            if (value != null)
            {
                if (value is decimal || value is long || value is ulong || value is int)
                    currentCell.SetCellValue((double)value);
                else
                    currentCell.SetCellValue(value);
            }
            if (styleIndex != 0 && styleIndex < _workbook.NumCellStyles)
                currentCell.CellStyle = _workbook.GetCellStyleAt((short)styleIndex);
            if (!string.IsNullOrEmpty(formula))
                currentCell.CellFormula = formula;
        }

        public void WriteTitle(string[] titles, int sheetIndex, int rowIndex, int cellIndex = 0, int styleIndex = 1)
        {
            CreateRow(sheetIndex, rowIndex);
            for (var i = 0; i < titles.Length; i++)
            {
                WriteValue(sheetIndex, rowIndex, cellIndex + i, titles[i], styleIndex);
            }
        }

        public int WriteProperty<T>(T entity, int sheetIndex, int rowIndex, int cellIndex = 0, int styleIndex = 2)
        {
            var thisIndex = cellIndex;
            var firstProperties = entity.GetProperties();
            foreach (var property in firstProperties)
            {
                var value = entity.GetValue(property);
                WriteValue(sheetIndex, rowIndex, thisIndex, value, styleIndex);
                thisIndex++;
            }
            return thisIndex;
        }

        public void WriteObject<T>(ICollection<T> entities, int sheetIndex, int rowIndex)
        {
            if (!entities.Any())
                return;
            var titles = typeof(T).GetProperties().Select(o => o.GetDescription()).ToArray();

            WriteTitle(titles, sheetIndex, rowIndex);
            rowIndex++;
            WriteEnumerable(entities, sheetIndex, rowIndex);
        }

        public void WriteEnumerable<T>(IEnumerable<T> entities, int sheetIndex, int rowIndex)
        {
            foreach (var entity in entities)
            {
                CreateRow(sheetIndex, rowIndex);
                WriteProperty(entity, sheetIndex, rowIndex);
                rowIndex++;
            }
        }

        public byte[] WriteStream()
        {
            var ms = new MemoryStream();
            _workbook.Write(ms);
            _workbook = null;
            return ms.ToArray();
        }

        public void WriteFile(string filePath)
        {
            using (var fs = new FileStream(filePath, FileMode.Create))
            {
                _workbook.Write(fs);
            }
            _workbook = null;
        }

        public void Read(int sheetIndex, int rowIndex, Action<dynamic[]> action)
        {
            var currentSheet = _workbook.GetSheetAt(sheetIndex);

            for (var i = rowIndex; i < currentSheet.LastRowNum; i++)
            {
                var currentRow = currentSheet.GetRow(i);
                var count = currentRow.LastCellNum;
                var value = new dynamic[count];
                for (var j = 0; j < count; j++)
                {
                    value[j] = currentRow.GetRowData(j);
                }
                action(value);
            }
        }

        public T ReadProperty<T>(int sheetIndex, int rowIndex)
        {
            var cellIndex = 0;
            var currentSheet = _workbook.GetSheetAt(sheetIndex);
            var currentRow = currentSheet.GetRow(rowIndex);

            var entity = Activator.CreateInstance(typeof(T));
            var properties = typeof(T).GetProperties();
            foreach (var propertyInfo in properties)
            {
                var value = currentRow.GetCell(cellIndex).GetCellData();
                value = Convert.ChangeType(value, propertyInfo.PropertyType);
                propertyInfo.SetValue(entity, value, null);
                cellIndex++;
            }
            return (T)entity;
        }

        public IEnumerable<T> ReadEnumerable<T>(int sheetIndex, int rowIndex)
        {
            var currentSheet = _workbook.GetSheetAt(sheetIndex);
            ICollection<T> results = new List<T>(currentSheet.LastRowNum - rowIndex);

            for (var i = rowIndex; i < currentSheet.LastRowNum; i++)
            {
                results.Add(ReadProperty<T>(sheetIndex, i));
            }
            return results;
        }

        public void SetColumnWidth(int sheetIndex, int firstColumn, int[] widths)
        {
            var currentSheet = _workbook.GetSheetAt(sheetIndex);
            for (var i = 0; i < widths.Length; i++)
            {
                currentSheet.SetColumnWidth(firstColumn + i, (widths[i] + 2) * 256);
            }
        }

        public void SetStyle(int sheetIndex, int firstRow, int lastRow, int firstColumn, int lastColumn, int styleIndex)
        {
            var currentSheet = _workbook.GetSheetAt(sheetIndex);

            if (styleIndex == 0)
            {
                currentSheet.AddMergedRegion(new CellRangeAddress(firstRow, lastRow, firstColumn, lastColumn));
            }
            else
            {
                var cellStyle = _workbook.GetCellStyleAt((short)styleIndex);
                for (var i = firstRow; i <= lastRow; i++)
                {
                    for (var j = firstColumn; j <= lastColumn; j++)
                    {
                        currentSheet.GetRow(i).GetCell(j).CellStyle = cellStyle;
                    }
                }
            }
        }
    }
}