using ExcelHelper.Exporter;
using ExcelHelper.Exporter.Attributes;
using ExcelHelper.Exporter.Dtos;
using ExcelHelper.Exporter.Enums;
using NPOI.HSSF.Util;
using System;
using System.Collections.Generic;
using System.IO;
using ExcelHelper.Common;
using Xunit;

namespace ExcelHelper.Test
{
    public class ExporterTest
    {
        protected readonly IExcelExporter Exporter;
        public ExporterTest()
        {
            Exporter = new DefaultExcelExporter();
        }
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
                    Id = Guid.NewGuid(),
                    Name = i % 6 == 1 ? null : ($"name{i}"),
                    Age = i,
                    Phone = i % 8 == 1 ? "adsf123" : $"{1}{new Random().Next(100, 999)}{1}{new Random().Next(100, 999)}{2}{new Random().Next(0, 9)}{3}",
                    Birthday = i % 13 == 1 ? default(DateTime?) : DateTime.Now.AddDays(i),
                    Money = Math.Round(new Random(i).NextDouble(), 2),
                    SchoolDate = DateTime.Now.AddDays(i + 1),
                });
                var grade = new ExportGrade { Code = $"����������������{i}", GradeName = $"{i}�꼶�꼶�꼶�꼶�꼶�꼶�꼶�꼶�꼶�꼶" };
                if (i % 11 == 1)
                    grade.SetCellStyle(nameof(grade.Code), new CellStyle() { FillForegroundColor = HSSFColor.Red.Index, FontColor = HSSFColor.Blue.Index });
                grades.Add(grade);
                schools.Add(new ExportSchool { Name = $"{i}��ѧУ", Address = $"ѧУ��ַ{i}", Price = Math.Round(new Random().NextDouble(), 2) });
            }

            var stream = new FileStream("D://Export.xlsx", FileMode.Create, FileAccess.Write);
            Exporter.Export(new ExportBook()
            {
                Ext = ExtEnum.Xlsx,
                Sheets = new List<BookSheet> {
                new BookSheet(){  SheetName="����", Data=students},
                new BookSheet(){   Data=grades},
                new BookSheet(){   Data=schools,Title=new  SheetTitle("ѧУ�б�",true,18,default,HorizontalAlignEnum.Center),  FilterColumn=new List<string>(){ "ѧУ����","Price" } },
                }
            }, stream);

            stream.Dispose();
        }

        [Fact]
        public void ExportMergeRow()
        {
            var orders = new List<Order>();
            for (int i = 0; i < 100; i++)
            {
                var index = new Random(i).Next(i + 10, i + 13);
                var orderNumber = $"����{index}";
                orders.Add(new Order()
                {
                    Buyer = $"�µ���{index}",
                    Price = Math.Round(new Random(i).NextDouble(), 2),
                    BuyQty = new Random(i).Next(1, 10),
                    ProductName = $"��Ʒ{i}",
                    OrderNumber = orderNumber,
                    OrderNum2 = orderNumber,
                    ExportPrimaryKey = orderNumber
                });
            }

            var stream = new FileStream("D://ExportMergeRow.xlsx", FileMode.Create, FileAccess.Write);
            Exporter.Export(new ExportBook()
            {
                Ext = ExtEnum.Xlsx,
                Sheets = new List<BookSheet> { new BookSheet() { SheetName = "�����б�", Data = orders } }
            }, stream);

            stream.Dispose();
        }

        [Fact]
        public void ExportByFormat()
        {
            var users = new List<User>();
            for (int i = 0; i < 100; i++)
            {
                users.Add(new User()
                {
                    IdCard = "3710812001000022" + i.ToString("00")
                });
            }


            var stream = new FileStream("D://ExportByFormat.xlsx", FileMode.Create, FileAccess.Write);
            var byt = Exporter.Export(new ExportBook()
            {
                Ext = ExtEnum.Xlsx,
                Sheets = new List<BookSheet> { new BookSheet() { Data = users } }
            });
            stream.Write(byt, 0, byt.Length);
            stream.Dispose();
        }


        public class User : SheetRow
        {

            [ColumnName("���֤")]
            public string IdCard { get; set; }

        }

        public class ExportStudent : SheetRow
        {
            [ColumnName("Id")]
            public Guid Id { get; set; }
            [ColumnName("����")]
            public string Name { get; set; }
            [ColumnName("����")]
            public int Age { get; set; }

            [ColumnName("����")]
            [StringFormatter("yyyy-MM-dd HH:mm:ss")]
            public DateTime? Birthday { get; set; }

            [ColumnStyle(FontName = "���Ĳ���")]
            [ColumnName("��ѧʱ��")]
            [StringFormatter("yyyy-MM-dd")]
            public DateTime SchoolDate { get; set; }

            [ColumnStyle(FontColor = 211, IsBold = true)]
            [ColumnName("�㻨Ǯ")]
            public double Money { get; set; }

            [ColumnName("�绰")]
            public string Phone { get; set; }
        }

        public class ExportSchool : SheetRow
        {
            [ColumnName("ѧУ����")]
            public string Name { get; set; }

            [ColumnName("ѧУ��ַ")]
            public string Address { get; set; }

            [ColumnName("ѧ��")]
            public double Price { get; set; }
        }

        public class ExportGrade : SheetRow
        {
            [ColumnWidth(0, 10000)]
            [HeaderStyle(true)]
            [ColumnName("�꼶����")]
            public string GradeName { get; set; }

            [ColumnWidth(0, 5000)]
            [HeaderStyle(true, FontColor = HSSFColor.Blue.Index)]
            [ColumnStyle(WrapText = false, FontSize = 9)]
            [ColumnName("�꼶����")]
            public string Code { get; set; }
        }

        public class Order : SheetRow
        {
            [RowMerged]
            [HeaderStyle(isBold: true)]
            [ColumnStyle(isBold: true, FontColor = HSSFColor.Blue.Index, VerticalAlign = VerticalAlignmentEnum.Center, HorizontalAlign = HorizontalAlignEnum.Right)]
            [ColumnName("��������")]
            public string OrderNumber { get; set; }

            [RowMerged]
            [HeaderStyle(isBold: true, VerticalAlign = VerticalAlignmentEnum.Top, HorizontalAlign = HorizontalAlignEnum.Right)]
            [ColumnStyle(VerticalAlign = VerticalAlignmentEnum.Center, HorizontalAlign = HorizontalAlignEnum.Center)]
            [ColumnName("�µ���")]
            public string Buyer { get; set; }

            [ColumnName("��Ʒ")]
            [HeaderStyle(isBold: true)]
            public string ProductName { get; set; }


            [ColumnName("�۸�")]
            [HeaderStyle(isBold: true)]
            public double Price { get; set; }

            [ColumnName("��������")]
            [HeaderStyle(isBold: true)]
            public int BuyQty { get; set; }

            [ColumnName("��������˶�")]
            [HeaderStyle(isBold: true)]
            public string OrderNum2 { get; set; }
        }
    }
}