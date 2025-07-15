namespace VOR.Models
{
    /// <summary>
    /// Класс, описывающий единицу кабеля (один отрезок кабеля)
    /// </summary>
    public class Cable : ICable
    {
        /// <summary>
        /// Конструктор для создания класса
        /// </summary>
        public Cable() { }

        public Cable(string Object, string start, string end, string marking, string brand, int numberCores, double crossSection, double length)
        {
            this.Object = Object;
            Start = start;
            End = end;
            Marking = marking;
            Brand = brand;
            NumberCores = numberCores;
            CrossSection = crossSection;
            Length = length;
        }


        /// <summary>
        /// Объект, к которому относится кабель
        /// </summary>
        public string Object { get; set; }
        
        /// <summary>
        /// Начало кабельной линии
        /// </summary>
        public string Start { get; set; }
        
        /// <summary>
        /// Конец кабельной линии
        /// </summary>
        public string End { get; set; }
        
        /// <summary>
        /// Маркировка кабеля
        /// </summary>
        public string Marking { get; set; }
        
        /// <summary>
        /// Марка кабеля
        /// </summary>
        public string Brand { get; set; }
        
        /// <summary>
        /// Количество жил
        /// </summary>
        public int NumberCores { get; set; }
        
        /// <summary>
        /// Сечение кабеля
        /// </summary>
        public double CrossSection { get; set; }
        
        /// <summary>
        /// Длина кабеля
        /// </summary>
        public double Length { get; set; }
    }
}
