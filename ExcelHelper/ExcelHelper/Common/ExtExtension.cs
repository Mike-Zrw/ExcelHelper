using System;

namespace ExcelHelper.Common
{
    public static class ExtExtension
    {
        public static ExtEnum GetExt(this string fileName)
        {
            if (fileName.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase))
                return ExtEnum.Xlsx;
            if (fileName.EndsWith(".xls", StringComparison.OrdinalIgnoreCase))
                return ExtEnum.Xls;
            return default;
        }
    }
}
