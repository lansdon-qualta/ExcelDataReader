using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using ExcelDataReader.Core.NumberFormat;

namespace ExcelDataReader.Core.CsvFormat
{
    internal class CsvWorkbook : IWorkbook<CsvWorksheet>
    {
        public CsvWorkbook(Stream stream, Encoding encoding, char[] autodetectSeparators, int analyzeInitialCsvRows, int maxRowsPerSheet)
        {
            Stream = stream;
            Encoding = encoding;
            AutodetectSeparators = autodetectSeparators;
            AnalyzeInitialCsvRows = analyzeInitialCsvRows;
            MaxRowsPerSheet = maxRowsPerSheet;
        }

        public int ResultsCount => 1;

        public Stream Stream { get; }

        public Encoding Encoding { get; }

        public char[] AutodetectSeparators { get; }

        public int AnalyzeInitialCsvRows { get; }

        public int MaxRowsPerSheet { get; }

        public IEnumerable<CsvWorksheet> ReadWorksheets()
        {
            yield return new CsvWorksheet(Stream, Encoding, AutodetectSeparators, AnalyzeInitialCsvRows, MaxRowsPerSheet);
        }

        public NumberFormatString GetNumberFormatString(int index)
        {
            return null;
        }
    }
}
