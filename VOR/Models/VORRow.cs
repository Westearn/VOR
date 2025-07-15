using System.Collections;
using System.Collections.Generic;

namespace VOR.Models
{
    /// <summary>
    /// Класс для представления строки таблицы ВОР
    /// </summary>
    public class VORRow : IRow, IEnumerable<string>
    {
        /// <summary>
        /// Конструктор для создания класса
        /// </summary>
        public VORRow() { }

        /// <summary>
        /// Конструктор для создания класса
        /// </summary>
        public VORRow(int numberLevel, string name, string unit, string quantity, 
            string link, string calculation)
        {
            NumberLevel = numberLevel;
            Name = name;
            Unit = unit;
            Quantity = quantity;
            Link = link;
            Calculation = calculation;
        }

        /// <summary>
        /// № п/п
        /// 0 - 1 Уровень нумерации (1)
        /// 1 - 2 Уровень нумерации (1.1)
        /// 2 - 3 Уровень нумерации (1.1.1)
        /// </summary>
        public int NumberLevel { get; set; }
        
        /// <summary>
        /// Наименование работ
        /// </summary>
        public string Name { get; set; }

        /// <summary>
        /// Единица измерения
        /// </summary>
        public string Unit { get; set; }

        /// <summary>
        /// Количество
        /// </summary>
        public string Quantity { get; set; }

        /// <summary>
        /// Ссылка на чертежи, спецификации
        /// </summary>
        public string Link { get; set; }

        /// <summary>
        /// Расчет объемов работ и расхода материалов
        /// </summary>
        public string Calculation { get; set; }

        /// <summary>
        /// Enumerator для класса представления строки таблицы ВОР
        /// Не включает параметр "NumberLevel"
        /// </summary>
        /// <returns></returns>
        public IEnumerator<string> GetEnumerator()
        {
            yield return Name;
            yield return Unit;
            yield return Quantity;
            yield return Link;
            yield return Calculation;
        }

        IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();
    }
}
