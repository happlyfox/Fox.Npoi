namespace Fox.Npoi.Core
{
    public interface IStyle
    {
        void SetColumnWidth(int sheetIndex, int firstColumn, int[] widths);

        void SetStyle(int sheetIndex, int firstRow, int lastRow, int firstColumn, int lastColumn, int styleIndex);
    }
}