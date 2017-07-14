using Fox.Npoi.Style;
using Fox.Npoi.Test.ExtStyle;
using Fox.Npoi.Test.Model;
using System;
using System.IO;
using System.Linq;

namespace Fox.Npoi.Test
{
    public class DataShow
    {
        private static readonly string appPath = AppDomain.CurrentDomain.BaseDirectory;

        /// <summary>
        /// 导出标准数据带标题
        /// </summary>
        public static void Basic()
        {
            var excel = new Excel();
            excel.CreateSheet("部门表");
            excel.WriteObject(DataUtils.GetDepartList(), 0, 0);
            excel.SetColumnWidth(0, 0, new[] { 5, 30 });
            excel.WriteFile(Path.Combine(appPath, "Depart.xlsx"));
        }

        /// <summary>
        /// 导出标准数据不带标题
        /// </summary>
        public static void Basic2()
        {
            var excel = new Excel();
            excel.CreateSheet("部门表");

            int rowIndex = 0;
            foreach (var dep in DataUtils.GetDepartList())
            {
                excel.CreateRow(0, rowIndex);
                excel.WriteProperty<Depart>(dep, 0, rowIndex);
                rowIndex++;
            }
            excel.SetColumnWidth(0, 0, new[] { 5, 30 });
            excel.WriteFile(Path.Combine(appPath, "Depart2.xlsx"));
        }

        /// <summary>
        /// 导出标准数据 默认样式【加边框线】
        /// </summary>
        public static void Basic3()
        {
            var excel = new Excel(new DefaultStyle());
            excel.CreateSheet("部门表");

            int rowIndex = 0;
            foreach (var dep in DataUtils.GetDepartList())
            {
                excel.CreateRow(0, rowIndex);
                excel.WriteProperty<Depart>(dep, 0, rowIndex);
                rowIndex++;
            }
            excel.SetColumnWidth(0, 0, new[] { 5, 30 });
            excel.WriteFile(Path.Combine(appPath, "Depart.3xlsx"));
        }

        /// <summary>
        /// 导出标准数据 自定义样式【部门名称列加粗】 【标题黄色底纹】
        /// </summary>
        public static void Basic4()
        {
            var excel = new Excel(new OwnStyle());
            excel.CreateSheet("部门表");

            int rowIndex = 0;
            excel.WriteTitle(new string[] { "id", "部门名称" }, 0, 0, 0);
            rowIndex++;
            foreach (var dep in DataUtils.GetDepartList())
            {
                excel.CreateRow(0, rowIndex);
                excel.WriteProperty<Depart>(dep, 0, rowIndex, 0, 3);
                rowIndex++;
            }
            excel.SetColumnWidth(0, 0, new[] { 5, 30 });
            excel.WriteFile(Path.Combine(appPath, "Depart4.xlsx"));
        }

        /// <summary>
        /// 导出复杂数据  单元格合并(边写边合并)
        /// </summary>
        public static void Basic5()
        {
            var excel = new Excel(new FormatStyle());
            excel.CreateSheet("人员表");
            excel.WriteTitle(new string[] { "id", "人员名称", "部门", "年龄", "地址" }, 0, 0);

            var user = DataUtils.GetUserList().OrderBy(u => u.DepId).ToList();

            int rowIndex = 1;
            int midDep = user[0].DepId;
            int startIndex = rowIndex;
            int endIndex = startIndex;

            for (int i = 0; i < user.Count; i++)
            {
                if (user[i].DepId != midDep)
                {
                    endIndex = rowIndex - 1;
                    //合并单元格
                    excel.SetStyle(0, startIndex, endIndex, 2, 2, 0);
                    startIndex = endIndex + 1;
                    midDep = user[i].DepId;
                    i--;
                    continue;
                }

                excel.CreateRow(0, rowIndex);
                excel.WriteProperty(user[i], 0, rowIndex);
                rowIndex++;
            }

            //合并单元格
            excel.SetStyle(0, startIndex, rowIndex - 1, 2, 2, 0);
            excel.SetColumnWidth(0, 0, new int[] { 10, 10, 15, 20, 20 });
            excel.WriteFile(Path.Combine(appPath, "Depart5.xlsx"));
        }

        /// <summary>
        /// 导出复杂数据  单元格合并(写完后合并)
        /// </summary>
        public static void Basic6()
        {
            var excel = new Excel(new FormatStyle());
            excel.CreateSheet("人员表");

            //排序
            var userList = DataUtils.GetUserList().OrderBy(u => u.DepId).ToList();
            excel.WriteObject(userList, 0, 0);

            int row = excel.InternalExcel.GetWorkbook().GetSheetAt(0).LastRowNum;
            var user = excel.ReadEnumerable<User>(0, 1);
            var disUser = user.Select(u => new { u.DepId }).Distinct();
            var collects = disUser.Select(u => new
            {
                u.DepId,
                depCount = user.Count(p => p.DepId == u.DepId)
            }).ToList();

            int startIndex = 1;
            int endIndex = startIndex;
            foreach (var collect in collects)
            {
                endIndex = startIndex + collect.depCount - 1;
                excel.SetStyle(0, startIndex, endIndex, 2, 2, 0);
                startIndex = endIndex + 1;
            }

            excel.SetColumnWidth(0, 0, new int[] { 10, 10, 15, 20, 20 });
            excel.WriteFile(Path.Combine(appPath, "Depart6.xlsx"));
        }
    }
}