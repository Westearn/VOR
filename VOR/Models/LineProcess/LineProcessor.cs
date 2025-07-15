using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Windows;
using NLog;

using VOR.Convert;
using VOR.Helpers;
using VOR.Helpers.Import;

namespace VOR.Models.LineProcess
{
    abstract class LineProcessor
    {
        public readonly Logger _logger = LogManager.GetCurrentClassLogger();
        public abstract bool CanProcess(SpecRow row);
        public abstract void Process(SpecRow row, out List<VORRow> row_out);
        public List<VORRow> Okrs(double kol_okrs, string calculation)
        {
            List<VORRow> row_out = new List<VORRow>();

            VORRow vORrow1 = new VORRow()
            {
                NumberLevel = 3,
                Name = "Очистка поверхности металлоконструкций до 2-й степени с обеспыливанием (в соответствии " +
                                "с ГОСТ 9.402-2004 и П4-06.01 ТТР-0002 версия 3.0, утвержденных приказом ПАО \"НК \"Роснефть\" " +
                                "от 31.12.2020 г. №185",
                Unit = "м^2",
                Quantity = Math.Round(kol_okrs, 2).ToString(),
                Link = "",
                Calculation = calculation
            };
            row_out.Add(vORrow1);

            VORRow vORrow2 = new VORRow()
            {
                NumberLevel = 3,
                Name = "Окраска металлоконструкций цинкнаполненной полиуретановой грунтовкой общей толщиной не менее 50 мкм",
                Unit = "м^2",
                Quantity = Math.Round(kol_okrs, 2).ToString(),
                Link = "",
                Calculation = calculation
            };
            row_out.Add(vORrow2);

            VORRow vORrow3 = new VORRow()
            {
                NumberLevel = 3,
                Name = "Окраска металлоконструкций полиуретановой эмалью общей толщиной не менее 150 мкм",
                Unit = "м^2",
                Quantity = Math.Round(kol_okrs, 2).ToString(),
                Link = "",
                Calculation = calculation
            };
            row_out.Add(vORrow3);

            return row_out;
        }
    }

    class AttributeProcessor : LineProcessor
    {
        public static string CurrentAttribute = "";
        public override bool CanProcess(SpecRow row) => row.Name.Trim() != "" && MainWindow.atr_kzh_dict.ContainsKey(row.Name.Trim());
        public override void Process(SpecRow row, out List<VORRow> row_out)
        {
            _logger.Info("Старт AttributeProcessor:" + row.Name);
            CurrentAttribute = row.Name.Trim();
            HeaderProcessor.CurrentHeader = "";
            row_out = new List<VORRow>();

            VORRow vORrow = new VORRow()
            {
                NumberLevel = 0,
                Name = row.Name,
                Unit = row.Unit,
                Quantity = row.Quantity,
                Link = "",
                Calculation = ""
            };
            row_out.Add(vORrow);
        }
    }

    class HeaderProcessor : LineProcessor
    {
        public static string CurrentHeader = "";
        public override bool CanProcess(SpecRow row) => row.Name.Trim() != "" && (row.Quantity.Trim() == "" || row.Unit.Trim() == "");
        public override void Process(SpecRow row, out List<VORRow> row_out)
        {
            _logger.Info("Старт HeaderProcessor:" + row.Name);
            CurrentHeader = row.Name.Trim();
            row_out = new List<VORRow>();

            VORRow vORrow = new VORRow()
            {
                NumberLevel = 1,
                Name = row.Name,
                Unit = row.Unit,
                Quantity = row.Quantity,
                Link = "",
                Calculation = ""
            };
            row_out.Add(vORrow);
        }
    }

    class InterBlockCablesProcessor : LineProcessor
    {
        private static List<(string atribute, string header, List<CableProducts> cables)> atr_cabels_dict = new List<(string atribute, string header, List<CableProducts>)>();
        public static ObservableCollection<InfoCableRow> NotFoundItems = new ObservableCollection<InfoCableRow>();
        public override bool CanProcess(SpecRow row) => row.Name.Contains("Кабель") && HeaderProcessor.CurrentHeader.Contains("Межблочная кабельная продукция");
        public override void Process(SpecRow row, out List<VORRow> row_out)
        {
            _logger.Info("Старт InterBlockCablesProcessor:" + row.Name);
            row_out = new List<VORRow>();

            // Создание строки для кабеля
            VORRow vORrow = new VORRow()
            {
                NumberLevel = 2,
                Name = row.Name,
                Unit = row.Unit,
                Quantity = row.Quantity,
                Link = "",
                Calculation = ""
            };
            row_out.Add(vORrow);

            if (!atr_cabels_dict.Any(x => x.atribute == AttributeProcessor.CurrentAttribute && x.header == HeaderProcessor.CurrentHeader))
            {
                MessageBox.Show($"Просьба загрузить кабельнотрубный журнал для: \n{HeaderProcessor.CurrentHeader}", "Межблочная кабельная продукция", MessageBoxButton.OK, MessageBoxImage.Information);

                var fileselector = new FileSelector();
                var filepath = fileselector.SelectFiles("Excel files (*.xlsx, *.xlsm)|*.xlsx; *.xlsm", false)[0];

                var cabelProductsImport = new CabelProductsImport();
                List<CableProducts> cabels = cabelProductsImport.Import(filepath);
                atr_cabels_dict.Add((AttributeProcessor.CurrentAttribute, HeaderProcessor.CurrentHeader, cabels));
            }

            var kZhToSmet = new KZhToSmet();
            var vORrows = kZhToSmet.Convert(row, atr_cabels_dict.First(x => x.atribute == AttributeProcessor.CurrentAttribute && x.header == HeaderProcessor.CurrentHeader).cables, out string info);

            if (vORrows == null)
            {
                NotFoundItems.Add(new InfoCableRow()
                {
                    Atribute = $"{AttributeProcessor.CurrentAttribute}: {HeaderProcessor.CurrentHeader}",
                    Name = row.Name,
                    Unit = row.Unit,
                    Quantity = row.Quantity,
                    Result = info
                });
                return;
            }

            row_out.AddRange(vORrows);
        }
    }
    
    class CabelProcessor : LineProcessor
    {
        private static Dictionary<string, List<CableProducts>> atr_cabels_dict = new Dictionary<string, List<CableProducts>>();
        public static ObservableCollection<InfoCableRow> NotFoundItems = new ObservableCollection<InfoCableRow>();
        public override bool CanProcess(SpecRow row) => row.Name.Contains("Кабель") && HeaderProcessor.CurrentHeader == "Кабельная продукция";
        public override void Process(SpecRow row, out List<VORRow> row_out)
        {
            _logger.Info("Старт CabelProcessor:" + row.Name);
            row_out = new List<VORRow>();

            // Создание строки для кабеля
            VORRow vORrow = new VORRow()
            {
                NumberLevel = 2,
                Name = row.Name,
                Unit = row.Unit,
                Quantity = row.Quantity,
                Link = "",
                Calculation = ""
            };
            row_out.Add(vORrow);

            if (string.IsNullOrEmpty(MainWindow.atr_kzh_dict[AttributeProcessor.CurrentAttribute]))
            {
                return;
            }

            if (atr_cabels_dict == null || !atr_cabels_dict.ContainsKey(AttributeProcessor.CurrentAttribute))
            {
                var cabelProductsImport = new CabelProductsImport();
                List<CableProducts> cabels = cabelProductsImport.Import(MainWindow.atr_kzh_dict[AttributeProcessor.CurrentAttribute]);
                atr_cabels_dict.Add(AttributeProcessor.CurrentAttribute, cabels);
            }

            var kZhToSmet = new KZhToSmet();
            var vORrows = kZhToSmet.Convert(row, atr_cabels_dict[AttributeProcessor.CurrentAttribute], out string info);

            if (vORrows == null)
            {
                NotFoundItems.Add(new InfoCableRow() 
                {
                    Atribute = AttributeProcessor.CurrentAttribute,
                    Name = row.Name,
                    Unit = row.Unit,
                    Quantity = row.Quantity,
                    Result = info 
                });
                return;
            }
            row_out.AddRange(vORrows);
        }
    }
    
    class FLCabelProcessor : LineProcessor
    {
        public static ObservableCollection<InfoCableRow> NotFoundItems = new ObservableCollection<InfoCableRow>();
        public override bool CanProcess(SpecRow row) => row.Name.Contains("Кабель") && HeaderProcessor.CurrentHeader == "Прожекторное освещение";
        public override void Process(SpecRow row, out List<VORRow> row_out)
        {
            _logger.Info("Старт CabelProcessor:" + row.Name);
            row_out = new List<VORRow>();

            // Создание строки для кабеля
            VORRow vORrow = new VORRow()
            {
                NumberLevel = 2,
                Name = row.Name,
                Unit = row.Unit,
                Quantity = row.Quantity,
                Link = "",
                Calculation = ""
            };
            row_out.Add(vORrow);

            if (row.Type.Contains("КГ-ХЛ") || row.Name.Contains("КГ-ХЛ"))
            {
                var row_kon = new VORRow()
                {
                    NumberLevel = 3,
                    Name = $"Кабель до 35 кВ по установленным конструкциям и лоткам с креплением на поворотах и в конце трассы на высоте от 15 до 30 м, масса 1 м кабеля: до 1 кг",
                    Quantity = Math.Round(double.Parse(vORrow.Quantity) / 100, 2).ToString(),
                    Unit = "100 м"
                };
                row_out.Add(row_kon);

                var row_zad = new VORRow()
                {
                    NumberLevel = 3,
                    Name = $"Заделка концевая сухая для 3-5-жильного кабеля с пластмассовой и резиновой изоляцией напряжением: до 1 кВ, сечение одной жилы до 35 мм2",
                    Quantity = Math.Round(double.Parse(vORrow.Quantity) * 2 / 1.5, 0).ToString(),
                    Unit = "шт."
                };
                row_out.Add(row_zad);
            }
            else if (row.Type.Contains("ВВГ") || row.Name.Contains("ВВГ"))
            {
                var row_trub = new VORRow()
                {
                    NumberLevel = 3,
                    Name = $"Кабель до 35 кВ в проложенных трубах, блоках и коробах, масса 1 м кабеля: до 1 кг",
                    Quantity = Math.Round(double.Parse(vORrow.Quantity) / 100, 2).ToString(),
                    Unit = "100 м"
                };
                row_out.Add(row_trub);

                var row_zad = new VORRow()
                {
                    NumberLevel = 3,
                    Name = $"Заделка концевая сухая для 3-5-жильного кабеля с пластмассовой и резиновой изоляцией напряжением: до 1 кВ, сечение одной жилы до 35 мм2",
                    Quantity = Math.Round(double.Parse(vORrow.Quantity) * 2 / 5, 0).ToString(),
                    Unit = "шт."
                };
                row_out.Add(row_zad);
            }
        }
    }

    class KTPProcessor : LineProcessor
    {
        public override bool CanProcess(SpecRow row) => row.Name.Contains("КТП") || row.Name.Contains("трансформаторная подстанция");
        public override void Process(SpecRow row, out List<VORRow> row_out)
        {
            _logger.Info("Старт KTPProcessor");
            row_out = new List<VORRow>();

            VORRow vORrow = new VORRow()
            {
                NumberLevel = 2,
                Name = row.Name,
                Unit = row.Unit,
                Quantity = row.Quantity,
                Link = "Принципиальная схема питающей сети",
                Calculation = ""
            };
            row_out.Add(vORrow);

            /*if (Main.PNR_logic_var)
            {
                string kit = row.Name.Split(' ').FirstOrDefault(a => a.Contains("КТПБ") || a.Contains("КТПЛП"));
                // "1ХКТПБ042М1010002С06Э-Ж1К1Н1П1Р1С1Т3У2"
                var kit_split = kit.Split('-');
                // [ "1ХКТПБ042М1010002С06Э", "Ж1К1Н1П1Р1С1Т3У2" ]
                var kit1 = kit_split[0];
                // "1ХКТПБ042М1010002С06Э"
                var kol_tr = kit1[0];
                // "1"
            }*/
        }
    }

    class NKUProcessor : LineProcessor
    {
        public override bool CanProcess(SpecRow row) => row.Name.Contains("НКУ") || row.Name.Contains("Низковольтное комплектное устройство");
        public override void Process(SpecRow row, out List<VORRow> row_out)
        {
            _logger.Info("Старт NKUProcessor");
            row_out = new List<VORRow>();

            VORRow vORrow = new VORRow()
            {
                NumberLevel = 2,
                Name = row.Name,
                Unit = row.Unit,
                Quantity = row.Quantity,
                Link = "Принципиальная схема распределительной сети",
                Calculation = ""
            };
            row_out.Add(vORrow);
        }
    }

    class PipeProcessor : LineProcessor
    {
        public override bool CanProcess(SpecRow row) => row.Name.Contains("Труба");
        public override void Process(SpecRow row, out List<VORRow> row_out)
        {
            _logger.Info("Старт PipeProcessor");
            row_out = new List<VORRow>();

            VORRow vORrow = new VORRow()
            {
                NumberLevel = 2,
                Name = row.Name,
                Unit = row.Unit,
                Quantity = row.Quantity,
                Link = "",
                Calculation = ""
            };
            row_out.Add(vORrow);

            foreach (var pipe in Constants.WGpipeDict)
            {
                if (vORrow.Name.Contains(pipe.NomDiam + "x") || vORrow.Name.Contains(pipe.NomDiam + "х"))
                {
                    double kol_okrs = 0;
                    switch (vORrow.Unit.Trim())
                    {
                        case "м":
                            kol_okrs = double.Parse(vORrow.Quantity) * 3.1416 * pipe.RealDiam / 1000;
                            break;
                        case "кг":
                            kol_okrs = double.Parse(vORrow.Quantity) * 3.1416 * pipe.RealDiam / 1000 / pipe.Weight;
                            break;
                        case "т":
                            kol_okrs = double.Parse(vORrow.Quantity) * 3.1416 * pipe.RealDiam / pipe.Weight;
                            break;
                    }

                    if (kol_okrs == 0)
                    {
                        _logger.Warn($"Не получилось рассчитать АКЗ для трубы \"{row.Name}\". Проверьте указанную единицу измерения: м, кг, т");
                        break;
                    }

                    row_out.AddRange(Okrs(kol_okrs, "Длина трубы * Длина внешней окружности трубы, где Длина внешней окружности трубы = Пи * Внешний диаметр трубы"));
                    break;
                }
            }
        }
    }

    class NakProcessor : LineProcessor
    {
        public override bool CanProcess(SpecRow row) => row.Name.Contains("Наконечник");
        public override void Process(SpecRow row, out List<VORRow> row_out)
        {
            _logger.Info("Старт NakProcessor");
            row_out = new List<VORRow>();

            VORRow vORrow = new VORRow()
            {
                NumberLevel = 2,
                Name = row.Name + " " + row.Type,
                Unit = row.Unit,
                Quantity = row.Quantity,
                Link = "",
                Calculation = ""
            };
            row_out.Add(vORrow);
        }
    }

    class TubeProcessor : LineProcessor
    {
        public override bool CanProcess(SpecRow row) => row.Name.Contains("Трубка термоусаживаемая");
        public override void Process(SpecRow row, out List<VORRow> row_out)
        {
            _logger.Info("Старт TubeProcessor");
            row_out = new List<VORRow>();

            VORRow vORrow = new VORRow()
            {
                NumberLevel = 2,
                Name = row.Name + " " + row.Type,
                Unit = row.Unit,
                Quantity = row.Quantity,
                Link = "",
                Calculation = ""
            };
            row_out.Add(vORrow);
        }
    }

    class GruntProcessor : LineProcessor
    {
        public override bool CanProcess(SpecRow row) => row.Name.Contains("Круг 12") ||
                                                        row.Name.Contains("Сталь круглая диаметром 12 мм") ||
                                                        row.Name.Contains("Круг, оцинкованный 12");
        public override void Process(SpecRow row, out List<VORRow> row_out)
        {
            _logger.Info("Старт GruntProcessor");
            row_out = new List<VORRow>();

            VORRow vORrow = new VORRow()
            {
                NumberLevel = 2,
                Name = row.Name,
                Unit = row.Unit,
                Quantity = row.Quantity,
                Link = "Молниезащита. Заземление",
                Calculation = ""
            };
            row_out.Add(vORrow);

            double kol_grunt = 0;
            switch (vORrow.Unit.Trim())
            {
                case "м":
                    kol_grunt = double.Parse(vORrow.Quantity) * 0.5 * 0.5 / 1000;
                    break;
                case "кг":
                    kol_grunt = double.Parse(vORrow.Quantity) * 0.5 * 0.5 / 0.89 / 1000;
                    break;
                case "т":
                    kol_grunt = double.Parse(vORrow.Quantity) * 0.5 * 0.5 / 0.89;
                    break;
            }

            if (kol_grunt == 0)
            {
                _logger.Warn($"Не получилось рассчитать грунт для круга \"{row.Name}\". Проверьте указанную единицу измерения: м, кг, т");
                return;
            }


            VORRow vORrow1 = new VORRow()
            {
                NumberLevel = 3,
                Name = "Разработка грунта в траншеях экскаватором \"обратная лопата\" с ковшом вместимостью 0,5 (0,5-0,63)" +
                        "\u00A0м3, в отвал группа грунтов: 1",
                Unit = "1000 м^3 грунта",
                Quantity = Math.Round(kol_grunt, 3).ToString(),
                Link = "",
                Calculation = "Длина стального круга * Ширина траншеи (0,5\u00A0м) * Глубина траншеи (0,5\u00A0м)"
            };
            row_out.Add(vORrow1);

            VORRow vORrow2 = new VORRow()
            {
                NumberLevel = 3,
                Name = "Уплотнение грунта пневматическими трамбовками, группа грунтов: 1-2",
                Unit = "1000 м^3 грунта",
                Quantity = Math.Round(kol_grunt, 3).ToString(),
                Link = "",
                Calculation = "Длина стального круга * Ширина траншеи (0,5\u00A0м) * Глубина траншеи (0,5\u00A0м)"
            };
            row_out.Add(vORrow2);

            VORRow vORrow3 = new VORRow()
            {
                NumberLevel = 3,
                Name = "Засыпка траншей и котлованов с перемещением грунта до 5\u00A0м бульдозерами мощностью: 59\u00A0кВт (80\u00A0л.с.)," +
                        " группа грунтов 2",
                Unit = "1000 м^3 грунта",
                Quantity = Math.Round(kol_grunt, 3).ToString(),
                Link = "",
                Calculation = "Длина стального круга * Ширина траншеи (0,5\u00A0м) * Глубина траншеи (0,5\u00A0м)"
            };
            row_out.Add(vORrow3);
        }
    }

    class UgolokProcessor : LineProcessor
    {
        public override bool CanProcess(SpecRow row) => row.Name.Contains("Уголок") &&
                                                        (row.Name.Contains("50x50x5") ||
                                                        row.Name.Contains("50х50х5"));
        public override void Process(SpecRow row, out List<VORRow> row_out)
        {
            _logger.Info("Старт UgolokProcessor");
            row_out = new List<VORRow>();

            VORRow vORrow = new VORRow()
            {
                NumberLevel = 2,
                Name = row.Name,
                Unit = row.Unit,
                Quantity = row.Quantity,
                Link = "",
                Calculation = ""
            };
            row_out.Add(vORrow);

            // Удельный вес уголка 50x50x5 по ГОСТ 8509-93 = 3,77 кг/м
            // Площадь поперечного сечения уголка 50x50x5 по ГОСТ 8509-93 = 4,8 см^2
            // Периметр поперечного сечения уголка 50x50x5 в соответствии с параметрами ГОСТ 8509-93 (без учета радиусов скругления) = (50 + 45 + 5) * 2 = 200 мм
            // Расчет площади поверхности по формуле: масса / удельный вес = длина => длина * периметр поперечного сечения + две площади поперечного сечения = площадь поверхности
            double kol_okrs = 0;

            switch (vORrow.Unit.Trim())
            {
                case "т":
                    kol_okrs = double.Parse(vORrow.Quantity) / 3.77 * 200 + 2 * 4.8 / 10000;
                    break;
                case "кг":
                    kol_okrs = double.Parse(vORrow.Quantity) / 3.77 * 200 / 1000 + 2 * 4.8 / 10000;
                    break;
            }

            if (kol_okrs == 0)
            {
                _logger.Warn($"Не получилось рассчитать АКЗ для уголка 50х50х5 \"{row.Name}\". Проверьте указанную единицу измерения: кг, т");
                return;
            }

            row_out.AddRange(Okrs(kol_okrs, "Длина уголка * Периметр торцевой части уголка + 2 * Площадь поперечного сечения уголка"));
        }
    }

    class ListMKProcessor : LineProcessor
    {
        public override bool CanProcess(SpecRow row) => row.Name.Contains("Лист") &&
                                                        (row.Name.Contains("6x50x50") ||
                                                        row.Name.Contains("6х50х50"));
        public override void Process(SpecRow row, out List<VORRow> row_out)
        {
            _logger.Info("Старт ListMKProcessor");
            row_out = new List<VORRow>();

            VORRow vORrow = new VORRow()
            {
                NumberLevel = 2,
                Name = row.Name,
                Unit = row.Unit,
                Quantity = row.Quantity,
                Link = "",
                Calculation = ""
            };
            row_out.Add(vORrow);

            // Удельный вес листа 6x50x50 по ГОСТ 19903-2015 = 0,12 кг/шт.
            // Площадь поверхности листа 6x50x50 по ГОСТ 19903-2015 = (50 * 50 * 2 + 6 * 50 * 4) / 1000000 = 0,0062 м^2
            // Расчет площади поверхности по формуле: масса / удельный вес = количество => количество * площадь поверхности единицы = площадь поверхности
            double kol_okrs = 0;
            switch (vORrow.Unit.Trim())
            {
                case "т":
                    kol_okrs = double.Parse(vORrow.Quantity) * 0.0062 / 0.12 / 1000;
                    break;
                case "кг":
                    kol_okrs = double.Parse(vORrow.Quantity) * 0.0062 / 0.12;
                    break;
                case "шт.":
                    kol_okrs = double.Parse(vORrow.Quantity) * 0.0062;
                    break;
            }

            if (kol_okrs == 0)
            {
                _logger.Warn($"Не получилось рассчитать АКЗ для листа 6х50х50 \"{row.Name}\". Проверьте указанную единицу измерения: кг, т");
                return;
            }

            row_out.AddRange(Okrs(kol_okrs, "Площадь поверхности * Количество"));
        }
    }

    class ShvellerProcessor : LineProcessor
    {
        public override bool CanProcess(SpecRow row) => row.Name.Contains("Швеллер 10У");
        public override void Process(SpecRow row, out List<VORRow> row_out)
        {
            _logger.Info("Старт ShvellerProcessor");
            row_out = new List<VORRow>();

            VORRow vORrow = new VORRow()
            {
                NumberLevel = 2,
                Name = row.Name,
                Unit = row.Unit,
                Quantity = row.Quantity,
                Link = "",
                Calculation = ""
            };
            row_out.Add(vORrow);

            // Удельный вес швеллера 10У по ГОСТ 8240-97 = 8,59 кг/м
            // Площадь поперечного сечения швеллера 10У по ГОСТ 8240-97 = 10,9 см^2
            // Периметр поперечного сечения швеллера 10У в соответствии с параметрами ГОСТ 8240-97 (без учета радиусов скругления) = (100 + 46 + 40,5) * 2 = 373 мм
            // Расчет площади поверхности по формуле: масса / удельный вес = длина => длина * периметр поперечного сечения + две площади поперечного сечения = площадь поверхности
            double kol_okrs = 0;
            switch (vORrow.Unit.Trim())
            {
                case "т":
                    kol_okrs = double.Parse(vORrow.Quantity) / 8.59 * 373 + 2 * 10.9 / 10000;
                    break;
                case "кг":
                    kol_okrs = double.Parse(vORrow.Quantity) / 8.59 * 373 / 1000 + 2 * 10.9 / 10000;
                    break;
            }

            if (kol_okrs == 0)
            {
                _logger.Warn($"Не получилось рассчитать АКЗ для швеллера 10У \"{row.Name}\". Проверьте указанную единицу измерения: кг, т");
                return;
            }

            row_out.AddRange(Okrs(kol_okrs, "Длина швеллера * Периметр торцевой части швеллера + 2 * Площадь поперечного сечения швеллера"));
        }
    }

    class ListPSH150Processor : LineProcessor
    {
        public override bool CanProcess(SpecRow row) => row.Name.Contains("Лист") &&
                                                        (row.Name.Contains("1x250x150") ||
                                                        row.Name.Contains("1х250х150"));
        public override void Process(SpecRow row, out List<VORRow> row_out)
        {
            _logger.Info("Старт ListPSH150Processor");
            row_out = new List<VORRow>();

            VORRow vORrow = new VORRow()
            {
                NumberLevel = 2,
                Name = row.Name,
                Unit = row.Unit,
                Quantity = row.Quantity,
                Link = "",
                Calculation = ""
            };
            row_out.Add(vORrow);

            // Удельный вес листа 1x250x150 по ГОСТ 19903-2015 = 0,27 кг/шт.
            // Площадь поверхности листа 1x250x150 по ГОСТ 19903-2015 = (250 * 150 * 2 + 1 * 250 * 2 + 1 * 150 *2) / 1000000 = 0,0758 м^2
            // Расчет площади поверхности по формуле: масса / удельный вес = количество => количество * площадь поверхности единицы = площадь поверхности
            double kol_okrs = 0;
            switch (vORrow.Unit.Trim())
            {
                case "т":
                    kol_okrs = double.Parse(vORrow.Quantity) * 0.0758 / 0.27 / 1000;
                    break;
                case "кг":
                    kol_okrs = double.Parse(vORrow.Quantity) * 0.0758 / 0.27;
                    break;
                case "шт.":
                    kol_okrs = double.Parse(vORrow.Quantity) * 0.0758;
                    break;
            }

            if (kol_okrs == 0)
            {
                _logger.Warn($"Не получилось рассчитать АКЗ для листа 1х250х150 \"{row.Name}\". Проверьте указанную единицу измерения: кг, т");
                return;
            }

            row_out.AddRange(Okrs(kol_okrs, "Площадь поверхности * Количество"));
        }
    }

    class ListPSH250Processor : LineProcessor
    {
        public override bool CanProcess(SpecRow row) => row.Name.Contains("Лист") &&
                                                        (row.Name.Contains("1x250x250") ||
                                                        row.Name.Contains("1х250х250"));
        public override void Process(SpecRow row, out List<VORRow> row_out)
        {
            _logger.Info("Старт ListPSH250Processor");
            row_out = new List<VORRow>();

            VORRow vORrow = new VORRow()
            {
                NumberLevel = 2,
                Name = row.Name,
                Unit = row.Unit,
                Quantity = row.Quantity,
                Link = "",
                Calculation = ""
            };
            row_out.Add(vORrow);

            // Удельный вес листа 1x250x250 по ГОСТ 19903-2015 = 0,49 кг/шт.
            // Площадь поверхности листа 1x250x250 по ГОСТ 19903-2015 = (250 * 250 * 2 + 1 * 250 * 4) / 1000000 = 0,126 м^2
            // Расчет площади поверхности по формуле: масса / удельный вес = количество => количество * площадь поверхности единицы = площадь поверхности
            double kol_okrs = 0;
            switch (vORrow.Unit.Trim())
            {
                case "т":
                    kol_okrs = double.Parse(vORrow.Quantity) * 0.126 / 0.49 / 1000;
                    break;
                case "кг":
                    kol_okrs = double.Parse(vORrow.Quantity) * 0.126 / 0.49;
                    break;
                case "шт.":
                    kol_okrs = double.Parse(vORrow.Quantity) * 0.126;
                    break;
            }

            if (kol_okrs == 0)
            {
                _logger.Warn($"Не получилось рассчитать АКЗ для листа 1х250х250 \"{row.Name}\". Проверьте указанную единицу измерения: кг, т");
                return;
            }

            row_out.AddRange(Okrs(kol_okrs, "Площадь поверхности * Количество"));
        }
    }

    class SvetGolProcessor : LineProcessor
    {
        // Светильник головной обрабатываем таким образом, что его не попадало в ВОР
        // Для этого прокидываем пустой список
        public override bool CanProcess(SpecRow row) => row.Name.Contains("головной");
        public override void Process(SpecRow row, out List<VORRow> row_out)
        {
            _logger.Info("Старт SvetGolProcessor");
            row_out = new List<VORRow>();
        }
    }

    class CommonProcessor : LineProcessor
    {
        public override bool CanProcess(SpecRow row) => row.Name.Trim() != "" && row.Quantity.Trim() != "" && row.Unit.Trim() != "";
        public override void Process(SpecRow row, out List<VORRow> row_out)
        {
            _logger.Info("Старт CommonProcessor:" + row.Name);
            row_out = new List<VORRow>();

            VORRow vORrow = new VORRow()
            {
                NumberLevel = 2,
                Name = row.Name,
                Unit = row.Unit,
                Quantity = row.Quantity,
                Link = "",
                Calculation = ""
            };
            row_out.Add(vORrow);
        }
    }

    class LineProcessorFactory
    {
        private readonly List<LineProcessor> _processors = new List<LineProcessor>();

        public void RegisterProcessor(LineProcessor processor)
        {
            _processors.Add(processor);
        }

        public LineProcessor GetProcessor(SpecRow row)
        {
            return _processors.FirstOrDefault(p => p.CanProcess(row));
        }
    }
}