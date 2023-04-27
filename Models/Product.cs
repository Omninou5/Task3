using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Task3.Models
{
    internal class Product
    {
        public string Id { get; set; }
        public string Name { get; set; }

        /// <summary>
        /// Единица измерения
        /// </summary>
        public string Unit { get; set; }
        
        /// <summary>
        /// Цена за единицу товара
        /// </summary>
        public string Price { get; set; }
    }
}