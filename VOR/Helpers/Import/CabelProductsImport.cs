using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using ClosedXML.Excel;

using VOR.Models;

namespace VOR.Helpers.Import
{
    public class CabelProductsImport : IImporter<CableProducts>
    {
        /// <summary>
        /// Метод для выписывания кабелей
        /// </summary>
        public List<CableProducts> Import(string filePath)
        {
            List<CableProducts> cabelProductsList = new List<CableProducts>();
            using (XLWorkbook workbook = new XLWorkbook(filePath))
            {
                IXLWorksheet worksheet;
                try
                {
                    worksheet = workbook.Worksheet("Задание сметчикам на кабель");
                }
                catch (ArgumentException)
                {
                    worksheet = workbook.Worksheet("Задание сметчикам кабель");
                }
                var lastCellUsed = worksheet.Column(1).CellsUsed().LastOrDefault();
                var row = lastCellUsed.Address.RowNumber;

                for (int i = 5; i <= row; i++)
                {
                    var size = worksheet.Cell(i, "B").GetText().Split('x', 'х');

                    var lengthPipe = worksheet.Cell(i, "D").Value.ToString().Split('/');
                    var diameterPipe = worksheet.Cell(i, "E").Value.ToString().Split('/');

                    var lengthSleeve = worksheet.Cell(i, "F").Value.ToString().Split('/');
                    var diameterSleeve = worksheet.Cell(i, "G").Value.ToString().Split('/');

                    var cabelProducts = new CableProducts()
                    {
                        Brand = worksheet.Cell(i, "A").GetText(),
                        NumberCores = int.Parse(size[0]),
                        CrossSection = double.Parse(size[1].Replace('.', ',')),
                        Length = worksheet.Cell(i, "C").GetDouble(),
                        LengthPipe1 = double.Parse(lengthPipe[0]),
                        DiameterPipe1 = int.Parse(diameterPipe[0]),
                        LengthPipe2 = lengthPipe.Length == 2 ? double.Parse(lengthPipe[1]) : 0,
                        DiameterPipe2 = diameterPipe.Length == 2 ? int.Parse(diameterPipe[1]) : 0,
                        LengthPipe3 = lengthPipe.Length == 3 ? double.Parse(lengthPipe[2]) : 0,
                        DiameterPipe3 = diameterPipe.Length == 3 ? int.Parse(diameterPipe[2]) : 0,
                        LengthSleeve1 = double.Parse(lengthSleeve[0]),
                        DiameterSleeve1 = int.Parse(diameterSleeve[0]),
                        LengthSleeve2 = lengthSleeve.Length == 2 ? double.Parse(lengthSleeve[1]) : 0,
                        DiameterSleeve2 = diameterSleeve.Length == 2 ? int.Parse(diameterSleeve[1]) : 0,
                        LengthSleeve3 = lengthSleeve.Length == 3 ? double.Parse(lengthSleeve[2]) : 0,
                        DiameterSleeve3 = diameterSleeve.Length == 3 ? int.Parse(diameterSleeve[2]) : 0,
                        PoEstakade = worksheet.Cell(i, "I").GetDouble(),
                        VTransh = worksheet.Cell(i, "J").GetDouble(),
                        PoKonstr = worksheet.Cell(i, "K").GetDouble(),
                        PoStene = worksheet.Cell(i, "L").GetDouble(),
                        PoKonstrVTrube = worksheet.Cell(i, "M").GetDouble(),
                        VTranshVTrube = worksheet.Cell(i, "N").GetDouble(),
                        Zadelki = System.Convert.ToInt32(worksheet.Cell(i, "O").GetDouble()),
                        OtrCable = System.Convert.ToInt32(worksheet.Cell(i, "P").GetDouble()),
                        CountCableCores = System.Convert.ToInt32(worksheet.Cell(i, "Q").GetDouble())
                    };
                    cabelProductsList.Add(cabelProducts);
                }
            }
            return cabelProductsList;
        }
    }
}
