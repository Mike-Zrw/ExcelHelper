using ExcelHelper.Excel;
using ExcelHelper.Excel.Attributes;
using ExcelHelper.Excel.Exporter;
using NPOI.HSSF.Util;
using System;
using System.Collections.Generic;
using System.IO;
using Xunit;

namespace ExcelHelper.Test
{
    public class ExcelExporterTest
    {
        [Fact]
        public void Export()
        {
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
        }

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

        public class ExportSchool : ExportModel
        {
            [ColumnName("学校名称")]
            public string Name { get; set; }

            [ColumnName("学校地址")]
            public string Address { get; set; }

            [ColumnName("学费")]
            public double Price { get; set; }
        }

        public class ExportGrade : ExportModel
        {
            [HeaderStyle(true)]
            [ColumnNameAttribute("年级名称")]
            public string GradeName { get; set; }

            [HeaderStyle(true, FontColor = HSSFColor.Blue.Index)]
            [ColumnNameAttribute("年级编码")]
            public string Code { get; set; }
        }

    }
}