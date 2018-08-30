
using ClosedXML.Excel;
using CsvHelper;
using CsvHelper.Excel;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace ExcelWebProject
{

    public class CSVExcel
    {
        public CSVExcel(CSVExcelConfig configuracao)
        {
            Configuracao = configuracao;
        }

        public CSVExcelConfig Configuracao { get; set; }

        public T[] LerCSV<T>(Stream stream) where T : IExportavelCsvExcel
        {
            using (TextReader textReader = new StreamReader(stream))
            {
                using (var reader = new CsvReader(textReader))
                {
                    reader.Configuration.Delimiter = ";";
                    reader.Configuration.HasHeaderRecord = true;
                    return reader.GetRecords<T>().ToArray();
                }
            }
        }

        public byte[] GravarCSV<T>(IEnumerable<T> dados) where T : IExportavelCsvExcel
        {
            using (MemoryStream mem = new MemoryStream())
            {
                using (TextWriter textWriter = new StreamWriter(mem))
                {
                    var csv = new CsvWriter(textWriter);
                    csv.WriteHeader<T>();

                    foreach (var objeto in dados)
                        csv.WriteRecord(objeto);
                    return mem.ToArray();
                }
            }
        }

        public T[] LerExcel<T>(byte[] input) where T : IExportavelCsvExcel
        {
            using (Stream stream = new MemoryStream(input))
            using (XLWorkbook workbook = new XLWorkbook(stream))
            using (CsvReader csvReader = new CsvReader(new ExcelParser(workbook)))
            {
                csvReader.Configuration.TrimFields = Configuracao.TrimFields;
                return csvReader.GetRecords<T>().ToArray();
            }
        }

        public byte[] GravarExcel<T>(IEnumerable<T> dados) where T : IExportavelCsvExcel
        {
            using (MemoryStream mem = new MemoryStream())
            using (var workbook = new XLWorkbook(XLEventTracking.Disabled))
            {
                var worksheet = workbook.AddWorksheet("Plan1");
                using (var csvWriter = new CsvWriter(new ExcelSerializer(worksheet)))
                {
                    csvWriter.Configuration.TrimFields = Configuracao.TrimFields;
                    csvWriter.WriteHeader<T>();
                    foreach (var objeto in dados)
                        csvWriter.WriteRecord(objeto);
                }

                if (Configuracao.GravarEmFormatoTexto)
                    foreach (var row in worksheet.RowsUsed())
                        foreach (var cell in row.CellsUsed())
                            cell.DataType = XLCellValues.Text;

                workbook.SaveAs(mem);
                return mem.ToArray();
            }
        }
    }
}