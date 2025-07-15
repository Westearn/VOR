using System.Collections.Generic;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

using VOR.Models;

namespace VOR.Helpers.Import
{
    public class SpecImport : IImporter<SpecRow>
    {
        public List<SpecRow> Import(string filePath)
        {
            List<SpecRow> specRowList = new List<SpecRow>();

            using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(filePath, true))
            {
                var body = wordDoc.MainDocumentPart.Document.Body;

                // Количество столбцов в определяемой таблице
                var columnCount = 9;

                // Определение таблицы в спецификации с 9-ю столбцами и "Примечанием" в 9-ом столбце
                Table targetTable = body.Elements<Table>()
                    .FirstOrDefault(table => table.Elements<TableRow>().Any() 
                    && table.Elements<TableRow>().First().Elements<TableCell>().Count() == columnCount
                    && table.Elements<TableRow>().First().Elements<TableCell>().Last().InnerText.Contains("Примечание") == true);

                if (targetTable != null)
                {
                    var tableRows = targetTable.Elements<TableRow>().Skip(2);

                    foreach (var row in tableRows)
                    {
                        var cells = row.Elements<TableCell>().ToList();
                        var specRow = new SpecRow()
                        {
                            Name = GetTableCellText(cells[1]),
                            Type = cells[2].InnerText.Trim(),
                            Code = cells[3].InnerText.Trim(),
                            Supplier = cells[4].InnerText.Trim(),
                            Unit = cells[5].InnerText.Trim(),
                            Quantity = cells[6].InnerText.Trim(),
                            Weight = cells[7].InnerText.Trim(),
                            Note = cells[8].InnerText.Trim(),
                        };

                        specRowList.Add(specRow);
                    }
                }
            }

            return specRowList;
        }

        private string GetTableCellText(TableCell cell)
        {
            StringBuilder sb = new StringBuilder();

            foreach (var paragraph in cell.Elements<Paragraph>())
            {
                foreach (var run in paragraph.Elements<Run>())
                {
                    foreach (var text in run.Elements<Text>())
                    {
                        sb.Append(text.Text);
                    }
                }
                sb.AppendLine();
            }

            return sb.ToString().TrimEnd();
        }
    }
}
