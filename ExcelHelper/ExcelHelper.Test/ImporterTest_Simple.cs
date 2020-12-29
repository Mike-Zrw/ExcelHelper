using ExcelHelper.Attributes;
using ExcelHelper.Importer;
using ExcelHelper.Importer.Attributes;
using ExcelHelper.Importer.Dtos;
using System;
using System.IO;
using Xunit;
using Xunit.Abstractions;

namespace ExcelHelper.Test
{
    public class ImporterTest_Simple
    {
        protected readonly ITestOutputHelper Output;
        protected readonly IExcelImporter _importer;
        public ImporterTest_Simple(ITestOutputHelper output)
        {
            Output = output;
            _importer = new DefaultExcelImporter();
        }

        [Fact]
        public void TestImport()
        {
            var sheet = new BookSheet<ImportBalanceOfPayment>
            {
                UniqueValidationPrompt = "年月不可重复",
                HeaderRowIndex = 0,
                SheetIndex = 0
            };
            var bookmodel = new ImportBook().SetSheets(sheet);

            var inputFilePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Excels//Simple.xlsx");
            using var inputStrem = new FileStream(inputFilePath, FileMode.OpenOrCreate, FileAccess.Read);
            using var outStrem = new FileStream("E://Simple_Error.xlsx", FileMode.Create, FileAccess.Write);

            var importResult = _importer.ImportExcel(inputStrem, inputFilePath.GetExt(), bookmodel, outStrem);

            var success = importResult.ImportSuccess;
            var summaryErrorMsg = importResult.GetSummaryErrorMessage();
            var importData = importResult.Sheets[0].GetData<ImportBalanceOfPayment>();

            Output.WriteLine($"success:{success}");
            Output.WriteLine(summaryErrorMsg);
        }

    }
    public class ImportBalanceOfPayment : SheetRow
    {
        [ColumnUnique]
        [ColumnName("年月")]
        public DateTime? YearMonth { get; set; }

        [ColumnName("折旧")]
        public decimal? DepreciationCharges { get; set; }

        [ColumnName("运费")]
        public decimal? LogisticsCharges { get; set; }

        [ColumnName("人工成本")]
        public decimal? LabourCharges { get; set; }

        [ColumnName("其他支出")]
        public decimal? OtherCharges { get; set; }

        [ColumnName("维修费")]
        public decimal? MaintenanceCharges { get; set; }

        [ColumnName("差旅费")]
        public decimal? TravelCharges { get; set; }

        [ColumnName("业务招待费")]
        public decimal? EntertainCharges { get; set; }

        [ColumnName("发货量")]
        public int? DeliveringAmount { get; set; }

        [ColumnRequired("总收入不可为空")]
        [ColumnName("总收入")]
        public decimal? TotalRevenue { get; set; }
    }
}
