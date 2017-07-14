# Fox.Npoi
Fox.Npoi扩展

扩展主要基于三个接口

1.IRead 读接口

2.IWrite 写接口

3.IStyle 样式接口

  样式接口作用：
  实例化Excel文件时，描述当前Excel文件的样式，而不是在创建文件后，单独对每一种样式进行设置。
	做到一次设置，多次使用。
    

实例1

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
        
 实例2
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
