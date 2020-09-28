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
                grades.Add(new ExportGrade { Code = $"����{i}", GradeName = $"{i}�꼶" });
                schools.Add(new ExportSchool { Name = $"{i}��ѧУ", Address = $"ѧУ��ַ{i}", Price = Math.Round(new Random().NextDouble(), 2) });
            }
            var exporter = new DefaultExcelExporter();

            var stream = new FileStream("D://export.xlsx", FileMode.Create, FileAccess.Write);
            //var stream = new MemoryStream();
            exporter.Export(new ExportBook()
            {
                Ext =ExtEnum.XLSX,
                Sheets = new List<ExportSheet> {
                new ExportSheet(){  SheetName="����", Data=students},
                new ExportSheet(){   Data=grades},
                new ExportSheet(){   Data=schools,Title=new  ExportTitle("ѧУ�б�",true,18,default,Excel.Enums.HorizontalAlignEnum.Center),  FilterColumn=new List<string>(){ "ѧУ����","price" } },
                }
            }, stream);

            stream.Dispose();
        }

        public class ExportStudent : ExportModel
        {
            [ColumnNameAttribute("����")]
            public string Name { get; set; }
            [ColumnNameAttribute("����")]
            public int Age { get; set; }

            [ColumnNameAttribute("����")]
            [StringFormatter("yyyy-MM-dd HH:mm:ss")]
            public DateTime? Birthday { get; set; }

            [ColumnStyle(FontName = "���Ĳ���")]
            [ColumnNameAttribute("��ѧʱ��")]
            [StringFormatter("yyyy-MM-dd")]
            public DateTime SchoolDate { get; set; }

            [ColumnStyle(FontColor = 211, IsBold = true)]
            [ColumnNameAttribute("�㻨Ǯ")]
            public double Money { get; set; }

            [ColumnNameAttribute("�绰")]
            public string Phone { get; set; }
        }

        public class ExportSchool : ExportModel
        {
            [ColumnName("ѧУ����")]
            public string Name { get; set; }

            [ColumnName("ѧУ��ַ")]
            public string Address { get; set; }

            [ColumnName("ѧ��")]
            public double Price { get; set; }
        }

        public class ExportGrade : ExportModel
        {
            [HeaderStyle(true)]
            [ColumnNameAttribute("�꼶����")]
            public string GradeName { get; set; }

            [HeaderStyle(true, FontColor = HSSFColor.Blue.Index)]
            [ColumnNameAttribute("�꼶����")]
            public string Code { get; set; }
        }

    }
}