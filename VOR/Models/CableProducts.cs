namespace VOR.Models
{
    /// <summary>
    /// Класс, описывающий единицу кабельной продукции (один тип кабеля)
    /// </summary>
    public class CableProducts : ICable
    {
        /// <summary>
        /// Конструктор для создания класса
        /// </summary>
        public CableProducts() { }

        /// <summary>
        /// Конструктор для создания класса
        /// </summary>
        public CableProducts(string brand, int numberCores, double crossSection, double length, double lengthPipe1,
            int diameterPipe1, double lengthPipe2, int diameterPipe2, double lengthPipe3, int diameterPipe3,
            double lengthSleeve1, int diameterSleeve1, double lengthSleeve2, int diameterSleeve2, double lengthSleeve3, int diameterSleeve3,
            double poEstakade, double vTransh, double poKonstr, double poStene, double poKonstrVTrube, double vTranshVTrube, int zadelki,
            int otrCable, int countCableCores)
        {
            Brand = brand;
            NumberCores = numberCores;
            CrossSection = crossSection;
            Length = length;
            LengthPipe1 = lengthPipe1;
            DiameterPipe1 = diameterPipe1;
            LengthPipe2 = lengthPipe2;
            DiameterPipe2 = diameterPipe2;
            LengthPipe3 = lengthPipe3;
            DiameterPipe3 = diameterPipe3;
            LengthSleeve1 = lengthSleeve1;
            DiameterSleeve1 = diameterSleeve1;
            LengthSleeve2 = lengthSleeve2;
            DiameterSleeve2 = diameterSleeve2;
            LengthSleeve3 = lengthSleeve3;
            DiameterSleeve3 = diameterSleeve3;
            PoEstakade = poEstakade;
            VTransh = vTransh;
            PoKonstr = poKonstr;
            PoStene = poStene;
            PoKonstrVTrube = poKonstrVTrube;
            VTranshVTrube = vTranshVTrube;
            Zadelki = zadelki;
            OtrCable = otrCable;
            CountCableCores = countCableCores;
        }

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

        /// <summary>
        /// Длина трубы 1
        /// </summary>
        public double LengthPipe1 { get; set; } = 0;

        /// <summary>
        /// Диаметр трубы 1
        /// </summary>
        public int DiameterPipe1 { get; set; } = 0;

        /// <summary>
        /// Длина трубы 2
        /// </summary>
        public double LengthPipe2 { get; set; } = 0;

        /// <summary>
        /// Диаметр трубы 2
        /// </summary>
        public int DiameterPipe2 { get; set; } = 0;

        /// <summary>
        /// Длина трубы 3
        /// </summary>
        public double LengthPipe3 { get; set; } = 0;

        /// <summary>
        /// Диаметр трубы 3
        /// </summary>
        public int DiameterPipe3 { get; set; } = 0;

        /// <summary>
        /// Длина металлорукава 1
        /// </summary>
        public double LengthSleeve1 { get; set; } = 0;

        /// <summary>
        /// Диаметр металлорукава 1
        /// </summary>
        public int DiameterSleeve1 { get; set; } = 0;

        /// <summary>
        /// Длина металлорукава 2
        /// </summary>
        public double LengthSleeve2 { get; set; } = 0;

        /// <summary>
        /// Диаметр металлорукава 2
        /// </summary>
        public int DiameterSleeve2 { get; set; } = 0;

        /// <summary>
        /// Длина металлорукава 3
        /// </summary>
        public double LengthSleeve3 { get; set; } = 0;

        /// <summary>
        /// Диаметр металлорукава 3
        /// </summary>
        public int DiameterSleeve3 { get; set; } = 0;

        /// <summary>
        /// Длина кабеля по эстакаде
        /// </summary>
        public double PoEstakade { get; set; } = 0;

        /// <summary>
        /// Длина кабеля в траншее
        /// </summary>
        public double VTransh { get; set; } = 0;

        /// <summary>
        /// Длина кабеля по конструкциям
        /// </summary>
        public double PoKonstr { get; set; } = 0;

        /// <summary>
        /// Длина кабеля по стене
        /// </summary>
        public double PoStene { get; set; } = 0;

        /// <summary>
        /// Длина кабеля по конструкциям в трубе
        /// </summary>
        public double PoKonstrVTrube { get; set; } = 0;

        /// <summary>
        /// Длина кабеля в траншее в трубе
        /// </summary>
        public double VTranshVTrube { get; set; } = 0;

        /// <summary>
        /// Заделки
        /// </summary>
        public int Zadelki { get; set; } = 0;

        /// <summary>
        /// Количество отрезков кабеля
        /// </summary>
        public int OtrCable { get; set; } = 0;

        /// <summary>
        /// Количество подключаемых жил
        /// </summary>
        public int CountCableCores { get; set; } = 0;
    }
}
