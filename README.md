# 概述
基于NPOI的Excel导入导出类库。支持多sheet导入导出。导出字段过滤。特性配置导入验证，非空验证，唯一验证等
详细描述点击：[https://www.cnblogs.com/bluesummer/p/13744421.html#4694957](https://www.cnblogs.com/bluesummer/p/13744421.html#4694957 "https://www.cnblogs.com/bluesummer/p/13744421.html#4694957")

# 导出配置支持

- **HeaderStyleAttribute** ：表头样式，（颜色，字体，大小，加粗）
- **StringFormatterAttribute** ：格式化时间
- **ColumnWidthAttribute**： 列宽
- **ExportTitle**：导出标题
- **SheetName**
- **FilterColumn** ：导出指定列

# 导入配置支持
- **ColumnRegexAttribute**：正则判断
- **ColumnRequiredAttribute**：非空判断
- **ColumnUniqueAttribute**：唯一判断，（重复行）
- **UniqueValidationPrompt**：唯一验证提示
- **ImportSheet.ValidateHandler** : 业务逻辑判断
- **HeaderRowIndex**：列名所在行
- **ImportBook.DataErrorForegroundColor**：  错误前景色(红)
- **ImportBook.RepeatedErrorForegroundColor**： 重复前景色（黄）
- **ImportBook.DefaultForegroundColor**： 默认前景色（白）

##  导入结果说明
- **ImportSuccess** ：是否导入成功
- **GetSummaryErrorMessage()** : excel中的所有错误文字展示
- **GetNotDisplayErrorMessage()**: 无法在excel中标注的错误信息，比如sheet格式不正确，excel格式不正确等
- **outPutStream**： 错误的单元格添加样式及标注输出到文件流中

##导出示例

     		var students = new List<ExportStudent>();
            var grades = new List<ExportGrade>();
            var schools = new List<ExportSchool>();
            for (int i = 0; i < 100; i++)
            {
                students.Add(new ExportStudent
                {
                    Name = i % 6 == 1 ? null : ($"name{i}"),
                    Age = i,
                    Phone = i % 8 == 1 ? "adsf123" : $"{1}{new Random().Next(100, 999)}{1}{new Random().Next(100, 999)}{2}{new Random().Next(0, 9)}{3}",
                    Birthday = i % 13 == 1 ? default(DateTime?) : DateTime.Now.AddDays(i),
                    Money = Math.Round(new Random(i).NextDouble(), 2),
                    SchoolDate = DateTime.Now.AddDays(i + 1),
                });
                grades.Add(new ExportGrade { Code = $"编码{i}", GradeName = $"{i}年级" });
                schools.Add(new ExportSchool { Name = $"{i}号学校", Address = $"学校地址{i}", Price = Math.Round(new Random().NextDouble(), 2) });
            }
            var exporter = new DefaultExcelExporter();

            var stream = new FileStream("D://export.xlsx", FileMode.Create, FileAccess.Write);
            //var stream = new MemoryStream();
            exporter.Export(new ExportBook()
            {
                Ext =ExtEnum.XLSX,
                Sheets = new List<ExportSheet> {
                new ExportSheet(){  SheetName="测试", Data=students},
                new ExportSheet(){   Data=grades},
                new ExportSheet(){   Data=schools,Title=new  ExportTitle("学校列表",true,18,default,Excel.Enums.HorizontalAlignEnum.Center),  FilterColumn=new List<string>(){ "学校名称","price" } },
                }
            }, stream);

            stream.Dispose();
			
			public class ExportStudent : ExportModel
	        {
	            [ColumnNameAttribute("名字")]
	            public string Name { get; set; }
	            [ColumnNameAttribute("年龄")]
	            public int Age { get; set; }
	
	            [ColumnNameAttribute("生日")]
	            [StringFormatter("yyyy-MM-dd HH:mm:ss")]
	            public DateTime? Birthday { get; set; }
	
	            [ColumnStyle(FontName = "华文彩云")]
	            [ColumnNameAttribute("入学时间")]
	            [StringFormatter("yyyy-MM-dd")]
	            public DateTime SchoolDate { get; set; }
	
	            [ColumnStyle(FontColor = 211, IsBold = true)]
	            [ColumnNameAttribute("零花钱")]
	            public double Money { get; set; }
	
	            [ColumnNameAttribute("电话")]
	            public string Phone { get; set; }
	        }

			

##导入示例

 			var sheet1 = new ImportSheet<ImportStudent>
            {
                UniqueValidationPrompt = "零花钱不可重复",
                HeaderRowIndex = 0,
                SheetIndex = 0,
                ValidateHandler = (list) =>
                {
                    foreach (ImportStudent model in list)
                    {
                        if (model.IsValidated && model.Name == "name0")
                            model.SetError(nameof(model.Name), "名字不可为0");
                        if (model.IsValidated && model.Money < 0.5)
                            model.SetError(nameof(model.Money), "零花钱不可小于0.5");
                    }
                }
            };
            var sheet2 = new ImportSheet<ImportGrade>
            {
                HeaderRowIndex = 0,
                SheetIndex = 1
            };
            var sheet3 = new ImportSheet<ImportSchool>
            {
                HeaderRowIndex = 1,
                SheetIndex = 2,
                ValidateHandler = (list) =>
                {

                    foreach (var model in list)
                    {
                        if (model.Price > 0.5)
                            model.SetError(nameof(model.Price), "学费不可大于0.5");
                    }
                }
            };
            var import = new DefaultExcelImporter();
            using var inputStrem = new FileStream("D://export.xlsx", FileMode.OpenOrCreate, FileAccess.Read);
            using var outStrem = new FileStream("D://error.xlsx", FileMode.Create, FileAccess.Write);
            var bookmodel = new ImportBook();
            bookmodel.SetSheetModels(sheet1, sheet2, sheet3);
            var ret = import.ImportExcel(inputStrem, ExtEnum.XLSX, bookmodel, outStrem);
            var success = ret.ImportSuccess;
            var summaryErrorMsg = ret.GetSummaryErrorMessage();
            var notDisplayMsg = ret.GetNotDisplayErrorMessage();
            Output.WriteLine($"success:{success}");
            Output.WriteLine("summaryErrorMsg------------");
            Output.WriteLine(summaryErrorMsg);
            Output.WriteLine("notDisplayMsg------------");
            Output.WriteLine(notDisplayMsg);

			public class ImportStudent : ImportModel
		    {
		        [ColumnRequired("名字必填")]
		        [ColumnNameAttribute("名字")]
		        public string Name { get; set; }
		        [ColumnNameAttribute("年龄")]
		        public int Age { get; set; }
		
		        [ColumnRequired]
		        [ColumnNameAttribute("生日")]
		        public DateTime Birthday { get; set; }
		
		        [ColumnNameAttribute("入学时间")]
		        public DateTime SchoolDate { get; set; }
		
		        [ColumnUnique]
		        [ColumnNameAttribute("零花钱")]
		        public double Money { get; set; }
		
		        [ColumnNameAttribute("电话")]
		        [ColumnRegex(@"^[1]+[1-9]+\d{9}$", "电话格式不对")]
		        public string Phone { get; set; }
		    }
