using System.Collections.Generic;

namespace ExcelHelper.Exporter.Dtos
{
    public class BookSheet
    {
        public string SheetName { get; set; }

        public SheetTitle Title { get; set; }

        public IEnumerable<SheetRow> Data { get; set; }

        public List<string> FilterColumn { get; set; }
    }
}
