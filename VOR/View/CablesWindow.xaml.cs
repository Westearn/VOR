using System.Collections.ObjectModel;
using System.Windows;

using VOR.Models;

namespace VOR.View
{
    /// <summary>
    /// Логика взаимодействия для CablesWindow.xaml
    /// </summary>
    public partial class CablesWindow : Window
    {
        public ObservableCollection<InfoCableRow> TableData { get; set; }
        public CablesWindow(ObservableCollection<InfoCableRow> data)
        {
            InitializeComponent();
            TableData = data; // Используем переданные данные
            DataContext = this; // Устанавливаем контекст данных
        }

        private void Ok_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }
    }
}
