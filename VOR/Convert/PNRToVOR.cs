using System.Collections.Generic;
using System.Linq;
using ClosedXML.Excel;

using VOR.Models;

namespace VOR.Convert
{
    public class PNRToVOR
    {
        public List<VORRow> Convert(string filepath)
        {
            var vORRowList = new List<VORRow>();

            var searchValues = new List<string> { "Кол-во", "Кол.", "Общее кол-во", "Количество" };

            using (XLWorkbook workbook = new XLWorkbook(filepath))
            {
                foreach (var worksheet in workbook.Worksheets)
                {
                    var list = new List<(int row, int name, int ed, int qu)>();
                    var rowsUsed = worksheet.RowsUsed();

                    // Проверка на наличии информации на листе
                    if (rowsUsed == null || rowsUsed.Count() == 0)
                    {
                        continue;
                    }
                    bool lsr = false;
                    int ivbc = 0;
                    // Цикл сбора инофрмации о стандартных заголовоках
                    foreach (var row in worksheet.RowsUsed())
                    {
                        int name = 0;
                        int ed = 0;
                        int qu = 0;

                        foreach (var cell in row.Cells())
                        {
                            string cellValue = cell.GetValue<string>();

                            if (cellValue == "Наименование")
                            {
                                name = cell.Address.ColumnNumber;
                            }
                            else if (cellValue == "Наименование работ и затрат, единица измерения")
                            {
                                lsr = true;
                                name = cell.Address.ColumnNumber;
                                ed = cell.Address.ColumnNumber;
                            }
                            else if (cellValue == "Ед. изм.")
                            {
                                ed = cell.Address.ColumnNumber;
                            }
                            else if (searchValues.Any(value => cellValue.Contains(value)))
                            {
                                qu = cell.Address.ColumnNumber;
                            }
                            else if (cellValue == "ИТОГИ В БАЗИСНЫХ ЦЕНАХ")
                            {
                                ivbc = cell.Address.RowNumber;
                            }
                        }

                        if (name != 0 && ed != 0 && qu != 0)
                        {
                            list.Add((row.RowNumber(), name, ed, qu));
                            name = 0;
                            ed = 0;
                            qu = 0;
                        }
                    }

                    // Проверка на наличии хотя бы одного заголовка
                    if (!list.Any())
                    {
                        continue;
                    }

                    // Добавление в список последней используемой строки
                    if (lsr)
                    {
                        list.Add((ivbc, 0, 0, 0));
                    }
                    else
                    {
                        list.Add((worksheet.LastRowUsed().RowNumber() + 1, 0, 0, 0));
                    }

                    // Цикл для выписывания информации по найденным заголовкам
                    for (int i = 0; i < list.Count - 1; i++)
                    {
                        if (worksheet.Cell(list[i].row - 1, list[i].name - 2).GetValue<string>().Trim() != "" &&
                            worksheet.Cell(list[i].row - 1, list[i].name).GetValue<string>().Trim() == "")
                        {
                            VORRow vORRow = new VORRow()
                            {
                                NumberLevel = 1,
                                Name = worksheet.Cell(list[i].row - 1, list[i].name - 2).GetValue<string>().Trim(),
                                Unit = "",
                                Quantity = "",
                                Link = "",
                                Calculation = ""
                            };
                            vORRowList.Add(vORRow);
                        }
                        for (int n = list[i].row + 1; n < list[i+1].row; n++)
                        {
                            var zag = worksheet.Cell(n, list[i].name - 2).GetValue<string>().Trim();

                            string name;
                            string unit = null;

                            if (lsr)
                            {
                                var splittext = worksheet.Cell(n, list[i].name).GetValue<string>().Split('(');
                                name = splittext[0].Trim();
                                if (splittext.Length == 2)
                                {
                                    unit = splittext[1].Split(')')[0].Trim();
                                }
                                else if (splittext.Length == 3)
                                {
                                    name = name + " (" + splittext[1].Trim();
                                    unit = splittext[2].Split(')')[0].Trim();
                                }
                            }
                            else
                            {
                                name = worksheet.Cell(n, list[i].name).GetValue<string>().Trim();
                                unit = worksheet.Cell(n, list[i].ed).GetValue<string>().Trim();
                            }

                            var qu = worksheet.Cell(n, list[i].qu).GetValue<string>().Trim();

                            if (name == "3")
                            {
                                continue;
                            }
                            else if (name == "" && zag != "" && n != list[i+1].row - 1)
                            {
                                VORRow vORRow = new VORRow()
                                {
                                    NumberLevel = 1,
                                    Name = zag,
                                    Unit = "",
                                    Quantity = "",
                                    Link = "",
                                    Calculation = ""
                                };
                                vORRowList.Add(vORRow);
                            }
                            else if (qu != "" && qu != "0")
                            {
                                VORRow vORRow = new VORRow()
                                {
                                    NumberLevel = 2,
                                    Name = name,
                                    Unit = unit,
                                    Quantity = qu,
                                    Link = "",
                                    Calculation = ""
                                };
                                vORRowList.Add(vORRow);
                            }
                        }
                    }
                }
            }

            return vORRowList;
        }
    }
}
