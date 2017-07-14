using NPOI.HSSF.Util;
using NPOI.SS.UserModel;
using System.Globalization;

namespace Fox.Npoi.Ext
{
    public static class WorkBookExtensions
    {
        public static IFont SetFontHeight(this IFont font)
        {
            font.FontHeightInPoints = 12;
            return font;
        }

        public static ICellStyle SetBorder(this ICellStyle style)
        {
            style.BorderTop = BorderStyle.Thin;
            style.TopBorderColor = HSSFColor.Black.Index;
            style.BorderBottom = BorderStyle.Thin;
            style.BottomBorderColor = HSSFColor.Black.Index;
            style.BorderLeft = BorderStyle.Thin;
            style.LeftBorderColor = HSSFColor.Black.Index;
            style.BorderRight = BorderStyle.Thin;
            style.RightBorderColor = HSSFColor.Black.Index;
            return style;
        }

        public static ICellStyle SetBackgroundColor(this ICellStyle style, short color)
        {
            style.FillForegroundColor = color;
            style.FillPattern = FillPattern.SolidForeground;
            return style;
        }

        public static ICellStyle SetAlignment(this ICellStyle style)
        {
            style.Alignment = HorizontalAlignment.Center;
            return style;
        }

        public static dynamic GetCellData(this ICell cell)
        {
            switch (cell.CellType)
            {
                case CellType.Numeric:
                    return cell.NumericCellValue;

                case CellType.String:
                    return cell.StringCellValue;

                case CellType.Formula:
                    return cell.NumericCellValue;

                case CellType.Blank:
                    return "NULL";

                case CellType.Boolean:
                    return cell.BooleanCellValue;

                case CellType.Error:
                    return cell.ErrorCellValue.ToString(CultureInfo.InvariantCulture);
            }
            return null;
        }

        public static dynamic GetRowData(this IRow row, int cellIndex)
        {
            var cell = row.GetCell(cellIndex);
            if (cell == null)
            {
                return string.Empty;
            }
            return cell.GetCellData();
        }
    }
}