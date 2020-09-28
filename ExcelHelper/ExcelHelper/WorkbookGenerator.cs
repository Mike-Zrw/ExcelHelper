using System;
using System.IO;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace ExcelHelper.Excel
{
    public class WorkbookGenerator
    {
        public static IWorkbook GetIWorkbook(Stream fileStream, ExtEnum ext)
        {
            if (ext == ExtEnum.XLSX)
            {
                return new XSSFWorkbook(fileStream);
            }
            else if (ext == ExtEnum.XLS)
            {
                return new HSSFWorkbook(fileStream);
            }
            else
            {
                throw new Exception("excel格式无法解析");
            }
        }

        public static IWorkbook GetIWorkbook(ExtEnum ext)
        {
            if (ext == ExtEnum.XLSX)
            {
                return new XSSFWorkbook();
            }
            else if (ext == ExtEnum.XLS)
            {
                return new HSSFWorkbook();
            }
            else
            {
                throw new Exception("excel格式无法解析");
            }
        }
    }
}
