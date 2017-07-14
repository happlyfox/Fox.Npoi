using Fox.Npoi.Ext;
using Fox.Npoi.Style;
using NPOI.HSSF.Util;
using NPOI.SS.UserModel;

namespace Fox.Npoi.Test.ExtStyle
{
    public class OwnStyle : DefaultStyle
    {
        public override void RegisterTitleStyle(IWorkbook workbook, ICellStyle cellStyle)
        {
            cellStyle.SetBorder();
            cellStyle.SetBackgroundColor(HSSFColor.Yellow.Index);
            cellStyle.SetFont(workbook.GetFontAt(1));
        }

        public override void RegisterCustomStyle(IWorkbook workbook)
        {
            //styleIndex = 3 【字体加粗】
            ICellStyle cellStyle = workbook.CreateCellStyle();
            IFont font = workbook.CreateFont();
            font.Boldweight = (short)FontBoldWeight.Bold;
            cellStyle.SetFont(font);
            RegisterContentStyle(workbook, cellStyle);
        }
    }
}