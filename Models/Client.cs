using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Task3.Models
{
    internal class Client
    {
        public string Id { get; set; }
        public string Name { get; set; }
        public string Address { get; set; }

        /// <summary>
        /// Контактное лицо (ФИО)
        /// </summary>
        public string Contact { get; set; }
    }
}
