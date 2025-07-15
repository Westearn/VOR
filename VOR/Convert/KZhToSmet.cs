using System;
using System.Collections.Generic;
using System.Linq;

using VOR.Models;
using VOR.Helpers;

namespace VOR.Convert
{
    public class KZhToSmet
    {
        public List<VORRow> Convert(SpecRow specRow, List<CableProducts> cableProducts, out string Info)
        {
            var vORRow = new List<VORRow>();
            Info = null;
            var quantity = double.Parse(specRow.Quantity.Trim());
            var specName = specRow.Name.NormalizeString();
            var specType = specRow.Type.NormalizeString();

            var cableProduct = cableProducts.FirstOrDefault(a => (specType.Contains(a.Brand.NormalizeString()) || specName.Contains(a.Brand.NormalizeString())) &&
                                                                 (specName.Contains($"{a.NumberCores}x{a.CrossSection}") || specName.Contains($"{a.NumberCores}х{a.CrossSection}")) && 
                                                                 quantity - a.Length >= 0 &&
                                                                 (quantity - a.Length) / a.Length < 1.03);

            if (cableProduct == default)
            {
                cableProduct = cableProducts.FirstOrDefault(a => specType.Contains(a.Brand.Replace("кВ", "").Trim().NormalizeString()) &&
                                                                 (specName.Contains($"{a.NumberCores}x{a.CrossSection}") || specName.Contains($"{a.NumberCores}х{a.CrossSection}")) &&
                                                                 quantity - a.Length >= 0 &&
                                                                 (quantity - a.Length) / a.Length < 1.03);
            }

            if (cableProduct == default)
            {
                var _cableProduct = cableProducts.FirstOrDefault(a => specType.Contains(a.Brand.NormalizeString()) &&
                                                                      (specName.Contains($"{a.NumberCores}x{a.CrossSection}") || specName.Contains($"{a.NumberCores}х{a.CrossSection}")));
                if (_cableProduct == default)
                {
                    Info = "Кабель не найден в КЖ";
                    return null;
                }

                Info = $"Длина кабеля в КЖ ({Math.Round(_cableProduct.Length, 0)}) не соответствует длине в спецификации";
                return null;
            }

            // Получение удельной массы кабеля
            /*var standardCable = Constants.cableWeightDictionary.FirstOrDefault(a => cableProduct.Brand.Contains(a.Key.Brand)
                                                        && a.Key.NumberCores == cableProduct.NumberCores
                                                        && a.Key.CrossSection == cableProduct.CrossSection).Value;*/

            var key = (cableProduct.Brand.Split('-')[0].NormalizeString(), cableProduct.NumberCores, cableProduct.CrossSection);

            if (!Constants.cableWeightDictionary.TryGetValue(key, out double weight))
            {
                Info = "Кабель не найден в стандартном каталоге.\nОбратитесь в поддержку";
                return null;
            }

            // Число жил и сечение кабеля
            var numberCores = cableProduct.NumberCores;
            var crossSection = cableProduct.CrossSection;

            // Добавление расценки по эстакаде
            #region Расценки
            /*
                ФЕРм 08-02-151-01	Кабель до 35 кВ, прокладываемый по непроходным эстакадам, масса 1 м кабеля: до 3 кг
                ФЕРм 08-02-151-02	Кабель до 35 кВ, прокладываемый по непроходным эстакадам, масса 1 м кабеля: до 6 кг
                ФЕРм 08-02-151-03	Кабель до 35 кВ, прокладываемый по непроходным эстакадам, масса 1 м кабеля: до 13 кг
            */
            #endregion
            if (cableProduct.PoEstakade != 0)
            {
                // Определение массы для расценки по эстакаде
                var weight_est = weight < 3 ? 3 : weight < 6 ? 6 : 13;

                var row_est = new VORRow()
                {
                    NumberLevel = 3,
                    Name = $"Кабель до 35 кВ, прокладываемый по непроходным эстакадам, масса 1 м кабеля: до {weight_est} кг",
                    Quantity = Math.Round(cableProduct.PoEstakade / 100, 2).ToString(),
                    Unit = "100 м"
                };
                vORRow.Add(row_est);
            }

            // Определение массы для расценки по конструкциям
            var weight_kon = weight < 1 ? 1 : weight < 2 ? 2 : weight < 3 ? 3 : weight < 6 ? 6 : weight < 9 ? 9 : 13;

            // Добавление расценки по конструкциям
            #region Расценки
            /*
                ФЕРм 08-02-147-01	Кабель до 35 кВ по установленным конструкциям и лоткам с креплением на поворотах и в конце трассы, масса 1 м кабеля: до 1 кг
                ФЕРм 08-02-147-02	Кабель до 35 кВ по установленным конструкциям и лоткам с креплением на поворотах и в конце трассы, масса 1 м кабеля: до 2 кг
                ФЕРм 08-02-147-03	Кабель до 35 кВ по установленным конструкциям и лоткам с креплением на поворотах и в конце трассы, масса 1 м кабеля: до 3 кг
                ФЕРм 08-02-147-04	Кабель до 35 кВ по установленным конструкциям и лоткам с креплением на поворотах и в конце трассы, масса 1 м кабеля: до 6 кг
                ФЕРм 08-02-147-05	Кабель до 35 кВ по установленным конструкциям и лоткам с креплением на поворотах и в конце трассы, масса 1 м кабеля: до 9 кг
                ФЕРм 08-02-147-06	Кабель до 35 кВ по установленным конструкциям и лоткам с креплением на поворотах и в конце трассы, масса 1 м кабеля: до 13 кг
            */
            #endregion
            if (cableProduct.PoKonstr != 0)
            {
                var row_kon = new VORRow()
                {
                    NumberLevel = 3,
                    Name = $"Кабель до 35 кВ по установленным конструкциям и лоткам с креплением на поворотах и в конце трассы, масса 1 м кабеля: до {weight_kon} кг",
                    Quantity = Math.Round((cableProduct.PoKonstr + cableProduct.PoStene)/ 100, 2).ToString(),
                    Unit = "100 м"
                };
                vORRow.Add(row_kon);
            }

            // Добавление расценки по конструкциям в трубе
            #region Расценки
            /*
                ФЕРм 08-02-148-01	Кабель до 35 кВ в проложенных трубах, блоках и коробах, масса 1 м кабеля: до 1 кг
                ФЕРм 08-02-148-02	Кабель до 35 кВ в проложенных трубах, блоках и коробах, масса 1 м кабеля: до 2 кг
                ФЕРм 08-02-148-03	Кабель до 35 кВ в проложенных трубах, блоках и коробах, масса 1 м кабеля: до 3 кг
                ФЕРм 08-02-148-04	Кабель до 35 кВ в проложенных трубах, блоках и коробах, масса 1 м кабеля: до 6 кг
                ФЕРм 08-02-148-05	Кабель до 35 кВ в проложенных трубах, блоках и коробах, масса 1 м кабеля: до 9 кг
                ФЕРм 08-02-148-06	Кабель до 35 кВ в проложенных трубах, блоках и коробах, масса 1 м кабеля: до 13 кг
            */
            #endregion
            if (cableProduct.PoKonstrVTrube != 0 || cableProduct.LengthSleeve1 != 0)
            {
                var row_trub = new VORRow()
                {
                    NumberLevel = 3,
                    Name = $"Кабель до 35 кВ в проложенных трубах, блоках и коробах, масса 1 м кабеля: до {weight_kon} кг",
                    Quantity = Math.Round(cableProduct.PoKonstrVTrube / 100, 2).ToString(),
                    Unit = "100 м",
                    Calculation = "Длина трубы по конструкциям + Длина металлорукава"
                };
                vORRow.Add(row_trub);
            }

            // Добавление расценки в траншее
            #region Расценки
            /*
                ФЕРм 08-02-141-01	Кабель до 35 кВ в готовых траншеях без покрытий, масса 1 м: до 1 кг
                ФЕРм 08-02-141-02	Кабель до 35 кВ в готовых траншеях без покрытий, масса 1 м: до 2 кг
                ФЕРм 08-02-141-03	Кабель до 35 кВ в готовых траншеях без покрытий, масса 1 м: до 3 кг
                ФЕРм 08-02-141-04	Кабель до 35 кВ в готовых траншеях без покрытий, масса 1 м: до 6 кг
                ФЕРм 08-02-141-05	Кабель до 35 кВ в готовых траншеях без покрытий, масса 1 м: до 9 кг
                ФЕРм 08-02-141-06	Кабель до 35 кВ в готовых траншеях без покрытий, масса 1 м: до 13 кг
            */
            #endregion
            if (cableProduct.VTranshVTrube != 0 || cableProduct.VTransh != 0)
            {
                var row_transh = new VORRow()
                {
                    NumberLevel = 3,
                    Name = $"Кабель до 35 кВ в готовых траншеях без покрытий, масса 1 м: до {weight_kon} кг",
                    Quantity = Math.Round((cableProduct.VTranshVTrube + cableProduct.VTransh) / 100, 2).ToString(),
                    Unit = "100 м",
                };
                vORRow.Add(row_transh);

                var kol_grunt = (cableProduct.VTranshVTrube + cableProduct.VTransh) * 0.5 * 0.5 / 1000;

                VORRow vORrow1 = new VORRow()
                {
                    NumberLevel = 3,
                    Name = "Разработка грунта в траншеях экскаватором \"обратная лопата\" с ковшом вместимостью 0,5 (0,5-0,63)" +
                            "\u00A0м3, в отвал группа грунтов: 1",
                    Unit = "1000 м^3 грунта",
                    Quantity = Math.Round(kol_grunt, 3).ToString(),
                    Link = "",
                    Calculation = "Длина кабеля в траншее * Ширина траншеи (0,5\u00A0м) * Глубина траншеи (0,5\u00A0м)"
                };
                vORRow.Add(vORrow1);

                VORRow vORrow2 = new VORRow()
                {
                    NumberLevel = 3,
                    Name = "Уплотнение грунта пневматическими трамбовками, группа грунтов: 1-2",
                    Unit = "1000 м^3 грунта",
                    Quantity = Math.Round(kol_grunt, 3).ToString(),
                    Link = "",
                    Calculation = "Длина кабеля в траншее * Ширина траншеи (0,5\u00A0м) * Глубина траншеи (0,5\u00A0м)"
                };
                vORRow.Add(vORrow2);

                VORRow vORrow3 = new VORRow()
                {
                    NumberLevel = 3,
                    Name = "Засыпка траншей и котлованов с перемещением грунта до 5\u00A0м бульдозерами мощностью: 59\u00A0кВт (80\u00A0л.с.)," +
                            " группа грунтов 2",
                    Unit = "1000 м^3 грунта",
                    Quantity = Math.Round(kol_grunt, 3).ToString(),
                    Link = "",
                    Calculation = "Длина кабеля в траншее * Ширина траншеи (0,5\u00A0м) * Глубина траншеи (0,5\u00A0м)"
                };
                vORRow.Add(vORrow3);
            }
                
            // Определение типа кабеля
            var cableType = cableProduct.Brand.Contains("КВВГ") || cableProduct.Brand.Contains("КВБбШ") || cableProduct.Brand.Contains("ККО") ? "Контрольный" : cableProduct.Brand.Contains("КПБ") || cableProduct.Brand.Contains("КБП") || cableProduct.Brand.Contains("СК") ? "Высоковольтный" : "Силовой";

            // Добавление расценки для заделок
            if (cableProduct.Zadelki != 0)
            {
                var row_zad = new VORRow()
                {
                    NumberLevel = 3,
                    Quantity = cableProduct.Zadelki.ToString(),
                    Unit = "шт."
                };

                // Вспомогательные переменные для расценок
                double _crossSection;
                int _numberCores;

                switch (cableType)
                {
                    case "Силовой":
                        #region Расценки
                        /*
                            ФЕРм 08-02-158-14	Заделка концевая сухая для 3-5-жильного кабеля с пластмассовой и резиновой изоляцией напряжением: до 1 кВ, сечение одной жилы от 1,5 мм2 до 35 мм2
                            ФЕРм 08-02-158-15	Заделка концевая сухая для 3-5-жильного кабеля с пластмассовой и резиновой изоляцией напряжением: до 1 кВ, сечение одной жилы до 120 мм2
                            ФЕРм 08-02-158-16	Заделка концевая сухая для 3-5-жильного кабеля с пластмассовой и резиновой изоляцией напряжением: до 1 кВ, сечение одной жилы до 185 мм2
                            ФЕРм 08-02-158-17	Заделка концевая сухая для 3-5-жильного кабеля с пластмассовой и резиновой изоляцией напряжением: до 1 кВ, сечение одной жилы до 240 мм2
                        */
                        #endregion
                        _crossSection = crossSection <= 35 ? 35 : crossSection <= 120 ? 120 : crossSection <= 185 ? 185 : 240;
                        row_zad.Name = $"Заделка концевая сухая для 3-5-жильного кабеля с пластмассовой и резиновой изоляцией напряжением: до 1 кВ, сечение одной жилы до {_crossSection} мм2";
                        break;
                    case "Высоковольтный":
                        #region Расценки
                        /*
                            ФЕРм 08-02-158-18	Заделка концевая сухая для 3-5-жильного кабеля с пластмассовой и резиновой изоляцией напряжением: до 10 кВ, сечение одной жилы до 35 мм2
                            ФЕРм 08-02-158-19	Заделка концевая сухая для 3-5-жильного кабеля с пластмассовой и резиновой изоляцией напряжением: до 10 кВ, сечение одной жилы до 70 мм2
                            ФЕРм 08-02-158-20	Заделка концевая сухая для 3-5-жильного кабеля с пластмассовой и резиновой изоляцией напряжением: до 10 кВ, сечение одной жилы до 120 мм2
                            ФЕРм 08-02-158-21	Заделка концевая сухая для 3-5-жильного кабеля с пластмассовой и резиновой изоляцией напряжением: до 10 кВ, сечение одной жилы до 185 мм2
                            ФЕРм 08-02-158-22	Заделка концевая сухая для 3-5-жильного кабеля с пластмассовой и резиновой изоляцией напряжением: до 10 кВ, сечение одной жилы до 240 мм2
                        */
                        #endregion
                        _crossSection = crossSection <= 35 ? 35 : crossSection <= 70 ? 70 : crossSection <= 120 ? 120 : crossSection <= 185 ? 185 : 240;
                        row_zad.Name = $"Заделка концевая сухая для 3-5-жильного кабеля с пластмассовой и резиновой изоляцией напряжением: до 10 кВ, сечение одной жилы до {_crossSection} мм2";
                        break;
                    case "Контрольный":
                        #region Расценки
                        /*
                            ФЕРм 08-02-158-04	Заделка концевая сухая для контрольного кабеля сечением одной жилы: до 2,5 мм2, количество жил до 4
                            ФЕРм 08-02-158-05	Заделка концевая сухая для контрольного кабеля сечением одной жилы: до 2,5 мм2, количество жил до 7
                            ФЕРм 08-02-158-06	Заделка концевая сухая для контрольного кабеля сечением одной жилы: до 2,5 мм2, количество жил до 10
                            ФЕРм 08-02-158-07	Заделка концевая сухая для контрольного кабеля сечением одной жилы: до 2,5 мм2, количество жил до 14
                            ФЕРм 08-02-158-08	Заделка концевая сухая для контрольного кабеля сечением одной жилы: до 2,5 мм2, количество жил до 19
                            ФЕРм 08-02-158-09	Заделка концевая сухая для контрольного кабеля сечением одной жилы: до 2,5 мм2, количество жил до 27
                            ФЕРм 08-02-158-10	Заделка концевая сухая для контрольного кабеля сечением одной жилы: до 2,5 мм2, количество жил до 37
                            ФЕРм 08-02-158-11	Заделка концевая сухая для контрольного кабеля сечением одной жилы: до 6 мм2, количество жил до 4
                            ФЕРм 08-02-158-12	Заделка концевая сухая для контрольного кабеля сечением одной жилы: до 6 мм2, количество жил до 7
                            ФЕРм 08-02-158-13	Заделка концевая сухая для контрольного кабеля сечением одной жилы: до 6 мм2, количество жил до 10
                        */
                        #endregion
                        _crossSection = crossSection <= 2.5 ? 2.5 : 6;
                        _numberCores = numberCores <= 4 ? 4 : numberCores <= 7 ? 7 : numberCores <= 10 ? 10 : numberCores <= 14 ? 14 : numberCores <= 19 ? 19 : numberCores <= 27 ? 27 : 37;
                        row_zad.Name = $"Заделка концевая сухая для контрольного кабеля сечением одной жилы: до {_crossSection} мм2, количество жил до {_numberCores}";
                        break;
                }

                vORRow.Add(row_zad);
            }

            // Добавление расценки для количества подключаемых жил
            #region Расценки
            /*
                ФЕРм 08-02-144-01	Присоединение к зажимам жил проводов или кабелей сечением: до 2,5 мм2
                ФЕРм 08-02-144-02	Присоединение к зажимам жил проводов или кабелей сечением: до 6 мм2
                ФЕРм 08-02-144-03	Присоединение к зажимам жил проводов или кабелей сечением: до 16 мм2
                ФЕРм 08-02-144-04	Присоединение к зажимам жил проводов или кабелей сечением: до 35 мм2
                ФЕРм 08-02-144-05	Присоединение к зажимам жил проводов или кабелей сечением: до 70 мм2
                ФЕРм 08-02-144-06	Присоединение к зажимам жил проводов или кабелей сечением: до 150 мм2
                ФЕРм 08-02-144-07	Присоединение к зажимам жил проводов или кабелей сечением: до 240 мм2
            */
            #endregion
            if (cableProduct.CountCableCores != 0)
            {
                var _crossSection = crossSection <= 2.5 ? 2.5 : crossSection <= 6 ? 6 : crossSection <= 16 ? 16 : crossSection <= 35 ? 35 : crossSection <= 70 ? 70 : crossSection <= 150 ? 150 : 240;

                var row_ccc = new VORRow()
                {
                    NumberLevel = 3,
                    Name = $"Присоединение к зажимам жил проводов или кабелей сечением: до {_crossSection} мм2",
                    Quantity = cableProduct.CountCableCores.ToString(),
                    Unit = "шт."
                };

                vORRow.Add(row_ccc);
            }

            return vORRow;
        }
    }
}
