using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Task3.Models
{
    internal class Bid
    {
        public string Id { get; set; }
        public string ProductId { get; set; }
        public string ClientId { get; set; }
        public string Number { get; set; }

        /// <summary>
        /// Требуемое количество товара
        /// </summary>
        public string Quantity { get; set; }

        /// <summary>
        /// Дата размещения заявки
        /// </summary>
        public DateTime Date { get; set; }
    }
}
