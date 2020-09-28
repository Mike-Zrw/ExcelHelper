using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ExcelHelper.Importer.Models.Import;

namespace ExcelHelper.Importer.Models.Result
{
    public class ResultSheet<T> : ImportSheet<T>, IResultSheet
             where T : ImportModel
    {
        private string uniqueValidateErrorMessage;
        private List<List<RepeatRow>> repeatedRowIndexes;

        public bool IsValidated { get => this.SheetFormatErrorMessage == null && !(this.Data != null && this.Data.Any(p => !p.IsValidated)) && this.IsUniqueValidated; }

        /// <summary>
        /// 唯一验证是否成功
        /// </summary>
        public bool IsUniqueValidated { get; set; } = true;

        /// <summary>
        /// sheet格式错误信息
        /// </summary>
        public string SheetFormatErrorMessage { get; set; }

        public string UniqueValidateErrorMessage { get => this.uniqueValidateErrorMessage; }

        public IEnumerable<ImportModel> ErrorRows { get => this.Data?.Where(p => !p.IsValidated); }

        /// <summary>
        /// 数据重复的行
        /// </summary>
        public List<List<RepeatRow>> RepeatedRowIndexes { get => this.repeatedRowIndexes; }

        public List<T> Data { get; set; }

        public void SetData(IEnumerable<T> data)
        {
            this.Data = data.ToList();
        }

        public void Validate()
        {
            this.ValidateHandler?.Invoke(this.Data);
            this.IsUniqueValidated = this.UniqueValidate();
        }

        public string GetSummaryErrorMessage()
        {
            if (this.IsValidated)
            {
                return null;
            }

            if (!string.IsNullOrWhiteSpace(this.SheetFormatErrorMessage))
            {
                return this.SheetFormatErrorMessage;
            }

            if (this.Data == null)
            {
                this.SheetFormatErrorMessage = "未找到符合条件的sheet";
                return this.SheetFormatErrorMessage;
            }

            var errorData = this.Data.Where(m => !m.IsValidated);
            var errMsg = new StringBuilder();
            foreach (ImportModel item in errorData)
            {
                foreach (var err in item.ColumnIndexError)
                {
                    errMsg.Append($"第{item.RowIndex + 1}行{err.Key + 1}列,{err.Value}.{Environment.NewLine}");
                }
            }

            if (!this.IsUniqueValidated)
            {
                errMsg.Append(this.uniqueValidateErrorMessage);
            }

            return errMsg.ToString();
        }

        public bool UniqueValidate()
        {
            if (!this.NeedUniqueValidation)
            {
                return true;
            }

            if (!this.Data?.Any() ?? true)
            {
                return true;
            }

            this.repeatedRowIndexes = this.Data.GroupBy(m => m.UniqueSign).Where(m => m.Count() > 1).Select(m => m.Select(p => new RepeatRow
            {
                RowIndex = p.RowIndex,
                ColumnIndexes = p.UniqueColumnIndexes,
            }).ToList()).ToList();
            if (!this.repeatedRowIndexes.Any())
            {
                return true;
            }

            var msg = new StringBuilder();
            this.repeatedRowIndexes.ForEach(item =>
            {
                msg.Append($"第{string.Join(",", item.Select(m => m.RowIndex + 1))}行.{this.UniqueValidationPrompt}.{Environment.NewLine}");
            });
            this.uniqueValidateErrorMessage = msg.ToString();
            return false;
        }
    }
}
