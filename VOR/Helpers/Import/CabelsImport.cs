using System.Collections.Generic;
using ClosedXML.Excel;

using VOR.Models;

namespace VOR.Helpers.Import
{
    public class CabelsImport : IImporter<Cable>
    {
        /// <summary>
        /// Метод для выписывания кабелей
        /// </summary>
        public List<Cable> Import(string filePath)
        {
            List<Cable> cabelList = new List<Cable>();

            using (XLWorkbook workbook = new XLWorkbook(filePath))
            {
                var worksheet = workbook.Worksheet("Общая база");
                var row = worksheet.LastRowUsed().RowNumber();

                for (int i = 1; i <= row; i++)
                {
                    var size = worksheet.Cell(i + 4, "J").GetText().Split('x', 'х');

                    var cabel = new Cable()
                    {
                        Object = worksheet.Cell(i + 4, "A").GetText(),
                        Start = worksheet.Cell(i + 4, "B").GetText() + " " +
                                worksheet.Cell(i + 4, "C").GetText() + " " +
                                worksheet.Cell(i + 4, "D").GetText(),
                        End = worksheet.Cell(i + 4, "E").GetText() + " " +
                              worksheet.Cell(i + 4, "F").GetText() + " " +
                              worksheet.Cell(i + 4, "G").GetText(),
                        Marking = worksheet.Cell(i + 4, "H").GetText(),
                        Brand = worksheet.Cell(i + 4, "I").GetText(),
                        NumberCores = int.Parse(size[0]),
                        CrossSection = double.Parse(size[1]),
                        Length = worksheet.Cell(i + 4, "I").GetDouble()
                    };

                    cabelList.Add(cabel);
                }
            }

            return cabelList;
        }
    }
}
