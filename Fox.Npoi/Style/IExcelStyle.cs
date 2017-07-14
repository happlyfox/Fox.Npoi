using NPOI.SS.UserModel;

namespace Fox.Npoi.Style
{
    public interface IExcelStyle
    {
        void InitCellStyle(IWorkbook workbook);

        void RegisterFont(IFont font);

        void RegisterTitleStyle(IWorkbook workbook, ICellStyle cellStyle);

        void RegisterContentStyle(IWorkbook workbook, ICellStyle cellStyle);

        void RegisterCustomStyle(IWorkbook workbook);
    }

    public static class ExcelStyleUtil
    {
        public static void RegisterStyle(IExcelStyle moduleStyle, IWorkbook workbook)
        {
            moduleStyle.InitCellStyle(workbook);
            moduleStyle.RegisterCustomStyle(workbook);
        }
    }
}