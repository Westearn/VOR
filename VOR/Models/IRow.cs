using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace VOR.Models
{
    public interface IRow
    {
        /// <summary>
        /// Наименование и техническая характеристика
        /// </summary>
        string Name { get; set; }

        /// <summary>
        /// Единица измерения
        /// </summary>
        string Unit { get; set; }

        /// <summary>
        /// Количество
        /// </summary>
        string Quantity { get; set; }
    }
}
