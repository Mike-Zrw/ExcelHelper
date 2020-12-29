using System;

namespace ExcelHelper
{
    public static class ExtExtension
    {
        public static ExtEnum GetExt(this string fileName)
        {
            if (fileName.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase))
                return ExtEnum.XLSX;
            if (fileName.EndsWith(".xls", StringComparison.OrdinalIgnoreCase))
                return ExtEnum.XLS;
            return default;
        }
    }
}
