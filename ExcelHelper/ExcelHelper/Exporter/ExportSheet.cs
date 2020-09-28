using System.Collections.Generic;

namespace ExcelHelper.Exporter
{
    public class ExportSheet
    {
        public string SheetName { get; set; }

        public ExportTitle Title { get; set; }

        public IEnumerable<ExportModel> Data { get; set; }

        public List<string> FilterColumn { get; set; }
    }
}
