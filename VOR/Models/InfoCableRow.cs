namespace VOR.Models
{
    public class InfoCableRow : IRow
    {
        /// <summary>
        /// Конструктор для создания класса
        /// </summary>
        public InfoCableRow() { }

        /// <summary>
        /// Конструктор для создания класса
        /// </summary>
        public InfoCableRow(string atribute, string name, string unit, string quantity, string result)
        {
            Atribute = atribute;
            Name = name;
            Unit = unit;
            Quantity = quantity;
            Result = result;
        }

        /// <summary>
        /// Атрибут
        /// </summary>
        public string Atribute { get; set; }

        /// <summary>
        /// Наименование
        /// </summary>
        public string Name { get; set; }

        /// <summary>
        /// Единица измерения
        /// </summary>
        public string Unit { get; set; }

        /// <summary>
        /// Количество по спецификации
        /// </summary>
        public string Quantity { get; set; }
        /// <summary>
        /// Примечание
        /// </summary>
        public string Result { get; set; }
    }
}
