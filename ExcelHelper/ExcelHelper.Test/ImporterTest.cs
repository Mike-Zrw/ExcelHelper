using ExcelHelper.Importer;
using ExcelHelper.Importer.Attributes;
using System;
using System.IO;
using ExcelHelper.Common;
using ExcelHelper.Importer.Dto;
using Xunit;
using Xunit.Abstractions;

namespace ExcelHelper.Test
{

    public class ImporterTest
    {
        protected readonly ITestOutputHelper Output;
        protected readonly IExcelImporter _importer;
        public ImporterTest(ITestOutputHelper output)
        {
            Output = output;
            _importer = new DefaultExcelImporter();
        }

        [Fact]
        public void TestImport()
        {
            var sheet1 = new BookSheet<ImportStudent>
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
            var sheet2 = new BookSheet<ImportGrade>
            {
                HeaderRowIndex = 0,
                SheetIndex = 1
            };
            var sheet3 = new BookSheet<ImportSchool>
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
            var bookmodel = new ImportBook().SetSheets(sheet1, sheet2, sheet3);

            using var inputStream = new FileStream(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Excels//Export.xlsx"), FileMode.OpenOrCreate, FileAccess.Read);
            using var outStream = new FileStream("D://Export_Error.xlsx", FileMode.Create, FileAccess.Write);
            var importResult = _importer.ImportExcel(inputStream, ExtEnum.Xlsx, bookmodel, outStream);

            var success = importResult.ImportSuccess;
            var summaryErrorMsg = importResult.GetSummaryErrorMessage();
            var notDisplayMsg = importResult.GetNotDisplayErrorMessage();
            var sheet1Data = importResult.Sheets[0].GetData<ImportStudent>();
            Output.WriteLine($"success:{success}");
            Output.WriteLine("summaryErrorMsg------------");
            Output.WriteLine(summaryErrorMsg);
            Output.WriteLine("notDisplayMsg------------");
            Output.WriteLine(notDisplayMsg);
        }

    }

    public class ImportStudent : SheetRow
    {
        [ColumnName("Id")]
        public Guid Id { get; set; }

        [ColumnRequired("名字必填")]
        [ColumnName("名字")]
        public string Name { get; set; }
        [ColumnName("年龄")]
        public int Age { get; set; }

        [ColumnRequired]
        [ColumnName("生日")]
        public DateTime Birthday { get; set; }

        [ColumnName("入学时间")]
        public DateTime SchoolDate { get; set; }

        [ColumnUnique]
        [ColumnName("零花钱")]
        public double Money { get; set; }

        [ColumnName("电话")]
        [ColumnRegex(@"^[1]+[1-9]+\d{9}$", "电话格式不对")]
        public string Phone { get; set; }
    }

    public class ImportSchool : SheetRow
    {
        [ColumnName("学校名称")]
        public string Name { get; set; }

        [ColumnName("学校地址")]
        public string Address { get; set; }

        [ColumnName("学费")]
        public double Price { get; set; }
    }

    public class ImportGrade : SheetRow
    {
        [ColumnName("年级名称")]
        public string GradeName { get; set; }

        [ColumnName("年级编码")]
        public string Code { get; set; }
    }

}
