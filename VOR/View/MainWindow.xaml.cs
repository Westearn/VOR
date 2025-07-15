using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows;

using VOR.Helpers;
using VOR.Helpers.Import;
using VOR.Models;

namespace VOR
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private Task _task1 = null;

        private static List<SpecRow> specRows = new List<SpecRow>();
        private static List<string> spec_list = new List<string>();
        public static List<string> kzh_list { get; private set; } = new List<string>();
        public static List<string> pnr_list = new List<string>();
        // Переменная, у которой в Key записывается атрибут, а в Value - соответствующий кабельнотрубный журнал
        public static Dictionary<string, string> atr_kzh_dict { get; private set; } = new Dictionary<string, string>();

        private void Window_Closed(object sender, EventArgs e)
        {
            Properties.Settings.Default.Save();
        }

        private async void Spec_button_Click(object sender, RoutedEventArgs e)
        {
            spec_pb.Value = 30;
            spec_block.Text = "Загрузка";

            var fileSelector = new FileSelector();

            // Спецификация выбирается в единичном экземпляре
            spec_list = fileSelector.SelectFiles("Word files (*.docx)|*.docx", false);

            var _spec_list = spec_list.ToList();
            if (_spec_list.Count > 0)
            {
                for (int i = 0; i < _spec_list.Count; i++)
                {
                    _spec_list[i] = _spec_list[i].Split('\\').Last();
                }
                spec_box.Text = string.Join(", ", _spec_list);

                var specImport = new SpecImport();
                // Запускаем отдельную задачу для импорта данных из спецификации
                _task1 = Task.Run(() => specRows = specImport.Import(spec_list[0]));
                await _task1;

                spec_pb.Value = 100;
                spec_block.Text = "Готово!";

                if (go_pb.Value != 0)
                {
                    go_pb.Value = 0;
                    go_block.Text = "Готово!";
                }
            }
            else
            {
                spec_pb.Value = 0;
                spec_block.Text = "Ожидание";
            }
        }

        private void Kzh_button_Click(object sender, RoutedEventArgs e)
        {
            kzh_pb.Value = 30;
            kzh_block.Text = "Загрузка";

            var fileSelector = new FileSelector();
            kzh_list = fileSelector.SelectFiles("Excel files (*.xlsx, *.xlsm)|*.xlsx; *.xlsm");

            var _kzh_list = kzh_list.ToList();
            if (_kzh_list.Count > 0)
            {
                for (int i = 0; i < _kzh_list.Count; i++)
                {
                    _kzh_list[i] = _kzh_list[i].Split('\\').Last();
                }
                kzh_box.Text = string.Join(", ", _kzh_list);

                kzh_pb.Value = 100;
                kzh_block.Text = "Готово!";

                if (go_pb.Value != 0)
                {
                    go_pb.Value = 0;
                    go_block.Text = "Готово!";
                }
            }
            else
            {
                kzh_pb.Value = 0;
                kzh_block.Text = "Ожидание";
            }
        }

        private void Pnr_button_Click(object sender, RoutedEventArgs e)
        {
            pnr_pb.Value = 30;
            pnr_block.Text = "Загрузка";

            var fileSelector = new FileSelector();
            pnr_list = fileSelector.SelectFiles("Excel files (*.xlsx, *.xlsm)|*.xlsx; *.xlsm");

            var _pnr_list = pnr_list.ToList();
            if (_pnr_list.Count > 0)
            {
                for (int i = 0; i < _pnr_list.Count; i++)
                {
                    _pnr_list[i] = _pnr_list[i].Split('\\').Last();
                }
                pnr_box.Text = string.Join(", ", _pnr_list);

                pnr_pb.Value = 100;
                pnr_block.Text = "Готово!";

                if (go_pb.Value != 0)
                {
                    go_pb.Value = 0;
                    go_block.Text = "Готово!";
                }
            }
            else
            {
                pnr_pb.Value = 0;
                pnr_block.Text = "Ожидание";
            }
        }

        private void Go_button_Click(object sender, RoutedEventArgs e)
        {
            if (_task1 == null)
            {
                MessageBox.Show("Сначала необходимо загрузить спецификацию");
                return;
            }
            if (!_task1.IsCompleted)
            {
                MessageBox.Show("Дождитесь загрузки спецификации");
                return;
            }

            go_pb.Value = 30;
            go_block.Text = "Выполнение";

            // Из Box с атрибутами формируются ключи словаря из списка атрибутов
            atr_kzh_dict = atr_box.Text.Split(',').Select(s => s.Trim()).ToDictionary(x => x, y => "");

            string filePath = Path.Combine(Path.GetDirectoryName(spec_list[0]), "Шаблон_ВОР.docx");
            string shablon = @"\\tnn\PIR\base\TechInfo\ELT\15_РАСЧЕТЫ и СПРАВОЧНИКИ\ВОР\Шаблон\Шаблон_ВОР.docx";
            if (!File.Exists(filePath))
            {
                if (!File.Exists(shablon))
                {
                    MessageBox.Show(@"Отсутствует шаблон: \\tnn\PIR\base\TechInfo\ELT\15_РАСЧЕТЫ и СПРАВОЧНИКИ\ВОР\Шаблон\Шаблон_ВОР.docx");
                    return;
                }
                File.Copy(shablon, filePath);
            }

            // Проверка файлов на доступность
            var fileSelector = new FileSelector();
            var isFileLocked = fileSelector.IsFileLocked(filePath);   

            if (isFileLocked)
            {
                // Если файл с шаблоном заблокирован, то вывести оповещение
                string message = "Пожалуйста, закройте следующие файлы перед продолжением:\n" + "Шаблон_ВОР.docx";
                MessageBox.Show(message, "Ошибка доступа", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            // Проверка наличия атрибутов среди заголовков спецификации в единичном экземпляре
            StringBuilder stringBuilder = new StringBuilder();
            if (atr_kzh_dict.First().Key.Count() > 0)
            {
                foreach (string atr in atr_kzh_dict.Keys)
                {
                    int counter = 0;
                    foreach (SpecRow row in specRows)
                    {
                        if (atr == row.Name)
                        {
                            counter++;
                        }
                    }

                    if (counter == 0)
                    {
                        stringBuilder.AppendLine($"Атрибут {atr} не найден среди заголовков");
                    }
                    else if (counter > 1)
                    {
                        stringBuilder.AppendLine($"Атрибут {atr} имеет более одного совпадения");
                    }
                }

                if (!string.IsNullOrEmpty(stringBuilder.ToString()))
                {
                    MessageBoxResult result = MessageBox.Show(stringBuilder.ToString(), "Ошибка при определении атрибутов", MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }

                // Проверка наличия КЖ соответсвующие атрибутам в единичном экземпляре
                foreach (string atr in atr_kzh_dict.Keys.ToList())
                {
                    int counter = 0;
                    string actually_kzh = "";
                    // Создаём паттерн для проверки атрибутов, чтобы исключить совпадение "1 этап" и, например, "11 этап"
                    string pattern = $@"(?<!\d){Regex.Escape(atr)}";

                    foreach (string kzh in kzh_list)
                    {
                        bool isMatch = Regex.IsMatch(kzh.Split('\\').Last(), pattern);

                        if (isMatch)
                        {
                            counter++;
                            actually_kzh = kzh;
                        }
                    }

                    if (counter == 1)
                    {
                        // Записываем соответсвующие КЖ для каждого атрибута
                        atr_kzh_dict[atr] = actually_kzh;
                    }
                    else if (counter == 0)
                    {
                        stringBuilder.AppendLine($"Для атрибута {atr} не найден кабельнотрубный журнал");
                    }
                    else if (counter > 1)
                    {
                        stringBuilder.AppendLine($"Для атрибута {atr} найдено более одного кабельнотрубного журнала");
                    }
                }

                if (!string.IsNullOrEmpty(stringBuilder.ToString()))
                {
                    stringBuilder.AppendLine("Продолжить?");
                    MessageBoxResult result = MessageBox.Show(stringBuilder.ToString(), "Предупреждение", MessageBoxButton.YesNo);

                    if (result != MessageBoxResult.Yes)
                    {
                        return;
                    }
                }
            }
            else if (kzh_list.Count() == 1)
            {
                atr_kzh_dict.Remove("");
                atr_kzh_dict.Add("Объект проектирования", kzh_list[0]);
                specRows.Insert(0, new SpecRow()
                {
                    Name = "Объект проектирования"
                });
            }
            else if (kzh_list.Count() == 0)
            {
                atr_kzh_dict.Remove("");
                atr_kzh_dict.Add("Объект проектирования", null);
                specRows.Insert(0, new SpecRow()
                {
                    Name = "Объект проектирования"
                });
            }
            else if (kzh_list.Count() > 1)
            {
                MessageBoxResult result = MessageBox.Show("Загружено несколько кабельнотрубных журналов, при этом не указаны атрибуты", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }
            
            Main main = new Main();
            main.CreateVOR(filePath, specRows);

            go_pb.Value = 100;
            go_block.Text = "Готово!";
        }
    }
}
