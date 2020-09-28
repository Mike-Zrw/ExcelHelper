using ExcelHelper.Excel;
using ExcelHelper.Excel.Attributes;
using ExcelHelper.Excel.Importer;
using ExcelHelper.Excel.Importer.Models.Import;
using System;
using System.IO;
using Xunit;
using Xunit.Abstractions;

namespace ExcelHelper.Test
{

    public class ExcelImporterTest
    {
        protected readonly ITestOutputHelper Output;
        public ExcelImporterTest(ITestOutputHelper output)
        {
            Output = output;
        }

        [Fact]
        public void TestImport()
        {
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
        }

    }

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

    public class ImportSchool : ImportModel
    {
        [ColumnName("学校名称")]
        public string Name { get; set; }

        [ColumnName("学校地址")]
        public string Address { get; set; }

        [ColumnName("学费")]
        public double Price { get; set; }
    }

    public class ImportGrade : ImportModel
    {
        [ColumnNameAttribute("年级名称")]
        public string GradeName { get; set; }

        [ColumnNameAttribute("年级编码")]
        public string Code { get; set; }
    }

}
