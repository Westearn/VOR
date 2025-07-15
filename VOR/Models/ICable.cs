namespace VOR.Models
{
    public interface ICable
    {
        /// <summary>
        /// Марка кабеля
        /// </summary>
        string Brand { get; set; }

        /// <summary>
        /// Количество жил
        /// </summary>
        int NumberCores { get; set; }

        /// <summary>
        /// Сечение кабеля
        /// </summary>
        double CrossSection { get; set; }

        /// <summary>
        /// Длина кабеля
        /// </summary>
        double Length { get; set; }
    }
}
