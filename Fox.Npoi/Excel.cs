using Fox.Npoi.Core;
using Fox.Npoi.Style;
using System;
using System.Collections.Generic;

namespace Fox.Npoi
{
    public class Excel : IRead, IBasicWrite, IStyle
    {
        public readonly InternalExcel InternalExcel;

        public Excel(IExcelStyle excelStyle = null, string filePath = null)
        {
            InternalExcel = new InternalExcel(excelStyle, filePath);
        }

        public void CreateRow(int sheetIndex, int rowIndex)
        {
            InternalExcel.CreateRow(sheetIndex, rowIndex);
        }

        public void WriteValue(int sheetIndex, int rowIndex, int cellIndex, dynamic value, int styleIndex, string formula = null)
        {
            InternalExcel.WriteValue(sheetIndex, rowIndex, cellIndex, value, styleIndex, formula);
        }

        public void CreateSheet(string name)
        {
            InternalExcel.CreateSheet(name);
        }

        public byte[] WriteStream()
        {
            return InternalExcel.WriteStream();
        }

        public void WriteFile(string filePath)
        {
            InternalExcel.WriteFile(filePath);
        }

        public void Read(int sheetIndex, int rowIndex, Action<dynamic[]> action)
        {
            InternalExcel.Read(sheetIndex, rowIndex, action);
        }

        public T ReadProperty<T>(int sheetIndex, int rowIndex)
        {
            return InternalExcel.ReadProperty<T>(sheetIndex, rowIndex);
        }

        public IEnumerable<T> ReadEnumerable<T>(int sheetIndex, int rowIndex)
        {
            return InternalExcel.ReadEnumerable<T>(sheetIndex, rowIndex);
        }

        public void SetColumnWidth(int sheetIndex, int firstColumn, int[] widths)
        {
            InternalExcel.SetColumnWidth(sheetIndex, firstColumn, widths);
        }

        public void SetStyle(int sheetIndex, int firstRow, int lastRow, int firstColumn, int lastColumn, int styleIndex)
        {
            InternalExcel.SetStyle(sheetIndex, firstRow, lastRow, firstColumn, lastColumn, styleIndex);
        }

        public void WriteTitle(string[] titles, int sheetIndex, int rowIndex, int cellIndex = 0, int styleIndex = 1)
        {
            InternalExcel.WriteTitle(titles, sheetIndex, rowIndex, cellIndex, styleIndex);
        }

        public int WriteProperty<T>(T entity, int sheetIndex, int rowIndex, int cellIndex = 0, int styleIndex = 2)
        {
            return InternalExcel.WriteProperty<T>(entity, sheetIndex, rowIndex, cellIndex, styleIndex);
        }

        public void WriteEnumerable<T>(IEnumerable<T> entities, int sheetIndex, int rowIndex)
        {
            InternalExcel.WriteEnumerable<T>(entities, sheetIndex, rowIndex);
        }

        public void WriteObject<T>(ICollection<T> entities, int sheetIndex, int rowIndex)
        {
            InternalExcel.WriteObject<T>(entities, sheetIndex, rowIndex);
        }
    }
}