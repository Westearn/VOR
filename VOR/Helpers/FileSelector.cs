using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows;
using Microsoft.Win32;

namespace VOR.Helpers
{
    public class FileSelector
    {
        /// <summary>
        /// Показывает окно выбора файлов и проверяет их доступность
        /// </summary>
        /// <param name="filter">Фильтр для указания типа выбираемых файлов</param>
        /// <returns>Список выбранных файлов</returns>
        public List<string> SelectFiles(string filter, bool multiselect = true)
        {
            // Создание OpenFileDialog
            OpenFileDialog openFileDialog = new OpenFileDialog
            {
                Filter = filter,
                Multiselect = multiselect
            };

            // Запуск диалогового окна
            if (openFileDialog.ShowDialog() == true)
            {
                var selectedFiles = openFileDialog.FileNames;

                // Проверка файлов на доступность
                var lockedFiles = selectedFiles.Where(IsFileLocked).ToList();

                if(lockedFiles.Any())
                {
                    // Если есть заблокированные файлы, показать сообщение с их перечислением
                    string message = "Пожалуйста, закройте следующие файлы перед продолжением:\n" +
                                     string.Join("\n", lockedFiles);
                    MessageBox.Show(message, "Ошибка доступа", MessageBoxButton.OK, MessageBoxImage.Warning);
                    
                    return new List<string>();
                }

                return selectedFiles.ToList();
            }
            else
            {
                MessageBox.Show("Выбор файлов отменен", "Оповещение", MessageBoxButton.OK, MessageBoxImage.Information);
                return new List<string>();
            }
        }

        /// <summary>
        /// Проверяет, заблокирован ли файл
        /// </summary>
        /// <param name="filePath">Путь к файлу</param>
        /// <returns>True, если файл заблокирован, иначе False</returns>
        public bool IsFileLocked(string filePath)
        {
            try
            {
                using (var stream = File.Open(filePath, FileMode.Open, FileAccess.ReadWrite, FileShare.None))
                {
                    // Если файл удалось открыть с полным доступом, он не заблокирован
                    return false;
                }
            }
            catch (IOException)
            {
                // Если возникло исключение, файл заблокирован
                return true;
            }
        }
    }
}
