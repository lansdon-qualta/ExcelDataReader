using System.IO;
using System.Text;
using ExcelDataReader.Core.BinaryFormat;

namespace ExcelDataReader
{
    /// <summary>
    /// ExcelDataReader Class
    /// </summary>
    internal class ExcelBinaryReader : ExcelDataReader<XlsWorkbook, XlsWorksheet>
    {
        public ExcelBinaryReader(Stream stream, string password, Encoding fallbackEncoding, int maxRowsPerSheet) 
            : base(maxRowsPerSheet)
        {
            Workbook = new XlsWorkbook(stream, password, fallbackEncoding, maxRowsPerSheet);

            // By default, the data reader is positioned on the first result.
            Reset();
        }

        public override void Close()
        {
            base.Close();
            Workbook?.Stream?.Dispose();
            Workbook = null;
        }
    }
}
