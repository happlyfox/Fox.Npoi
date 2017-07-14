# Fox.Npoi扩展
扩展主要基于三个接口
1.IRead 读接口

2.IWrite 写接口

3.IStyle 样式接口

实例化Excel文件时，描述当前Excel文件的样式，而不是在创建文件后，单独对每一种样式进行设置。做到一次设置，多次使用。
    
# 模板
	public class Depart
	{
	[Description("部门id")]
	public int DepId { get; set; }

	[Description("部门名称")]
	public string DepName { get; set; }
	}


# 基本导出
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

# 基本导出+默认样式
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

# 基本导出+重写样式
	/// <summary>
	/// 重写默认样式
	/// </summary>
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

# 复杂导出
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
