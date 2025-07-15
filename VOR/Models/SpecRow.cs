namespace VOR.Models
{
    /// <summary>
    /// Класс для представления строки таблицы спецификации
    /// </summary>
    public class SpecRow : IRow
    {
        /// <summary>
        /// Конструктор для создания класса
        /// </summary>
        public SpecRow() { }

        /// <summary>
        /// Конструктор для создания класса
        /// </summary>
        public SpecRow(string name, string type, string code, string supplier,
            string unit, string quantity, string weight, string note)
        {
            Name = name;
            Type = type;
            Code = code;
            Supplier = supplier;
            Unit = unit;
            Quantity = quantity;
            Weight = weight;
            Note = note;
        }

        /// <summary>
        /// Наименование и техническая характеристика
        /// </summary>
        public string Name { get; set; }

        /// <summary>
        /// Тип, марка, обозначение документа, опросного листа
        /// </summary>
        public string Type { get; set; }

        /// <summary>
        /// Код продукции
        /// </summary>
        public string Code { get; set; }

        /// <summary>
        /// Поставщик
        /// </summary>
        public string Supplier { get; set; }

        /// <summary>
        /// Единица измерения
        /// </summary>
        public string Unit { get; set; }

        /// <summary>
        /// Количество
        /// </summary>
        public string Quantity { get; set; }

        /// <summary>
        /// Масса единицы, кг
        /// </summary>
        public string Weight { get; set; }

        /// <summary>
        /// Примечание
        /// </summary>
        public string Note { get; set; }
    }
}
