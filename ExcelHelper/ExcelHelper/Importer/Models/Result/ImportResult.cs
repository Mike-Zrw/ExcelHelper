using System;
using System.Collections.Generic;
using System.Linq;

namespace ExcelHelper.Excel.Importer.Models.Result
{
    public class ImportResult
    {
        public List<IResultSheet> Sheets { get; set; }

        public Exception Exception { get; private set; }

        /// <summary>
        /// Excel本身存在的错误
        /// </summary>
        public string BookFormatErrorMessage { get; private set; }

        public bool ImportSuccess
        {
            get
            {
                return string.IsNullOrWhiteSpace(this.BookFormatErrorMessage) && !this.Sheets.Any(p => !p.IsValidated);
            }
        }

        public string GetSummaryErrorMessage()
        {
            if (!string.IsNullOrWhiteSpace(this.BookFormatErrorMessage))
            {
                return this.BookFormatErrorMessage;
            }

            return string.Join(Environment.NewLine, this.Sheets
                                   .Where(p => !p.IsValidated)
                                   .Select(m => $"{(string.IsNullOrWhiteSpace(m.SheetName) ? $"索引{m.SheetIndex}" : m.SheetName)}:{Environment.NewLine}{m.GetSummaryErrorMessage()}"));
        }

        /// <summary>
        /// 无法展示在excel中的错误信息
        /// </summary>
        /// <returns>string</returns>
        public string GetNotDisplayErrorMessage()
        {
            if (!string.IsNullOrWhiteSpace(this.BookFormatErrorMessage))
            {
                return this.BookFormatErrorMessage;
            }

            return string.Join(Environment.NewLine, this.Sheets.Where(p => !string.IsNullOrWhiteSpace(p.SheetFormatErrorMessage))
                          .Select(m => $"{(string.IsNullOrWhiteSpace(m.SheetName) ? $"索引{m.SheetIndex}" : m.SheetName)}:{Environment.NewLine}{m.SheetFormatErrorMessage}"));
        }

        public void SetBookFormatErrorMessage(string error, Exception ex = null)
        {
            this.BookFormatErrorMessage = error;
            this.Exception = ex;
        }

        public void SetSheets(params IResultSheet[] args)
        {
            this.Sheets = args.ToList();
        }
    }
}
