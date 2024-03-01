using System.IO;
using ExcelDataReader.Core.OpenXmlFormat;

namespace ExcelDataReader
{
    internal class ExcelOpenXmlReader : ExcelDataReader<XlsxWorkbook, XlsxWorksheet>
    {
        public ExcelOpenXmlReader(Stream stream, int maxRowsPerSheet = 0)
            : base(maxRowsPerSheet)
        {
            MaxRowsPerSheet = maxRowsPerSheet;
            Document = new ZipWorker(stream);
            Workbook = new XlsxWorkbook(Document, maxRowsPerSheet);

            // By default, the data reader is positioned on the first result.
            Reset();
        }

        private ZipWorker Document { get; set; }

        public override void Close()
        {
            base.Close();

            Document?.Dispose();
            Workbook = null;
            Document = null;
        }
    }
}
