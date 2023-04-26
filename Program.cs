using ClosedXML.Excel;
using Task3.Models;

namespace Task3
{
    internal class Program
    {
        static void Main(string[] args)
        {
            string filePath = String.Empty;
            while (File.Exists(filePath) != true)
            {
                Console.WriteLine("Введите полный путь к файлу:");
                filePath = Console.ReadLine();
                if (File.Exists(filePath) != true)
                {
                    Console.WriteLine("Неверный путь к файлу!");
                }
                Console.WriteLine();
            }

            List<Product> products;
            List<Client> clients;
            List<Bid> bids;
            ExtractProductsClientsBids(filePath, out products, out clients, out bids);

            while (true)
            {
                Console.WriteLine();
                Console.WriteLine("Введите цифру:");
                Console.WriteLine("2 - Вывод информации о клиентах, заказавших определенный товар.");
                Console.WriteLine("3 - Изменение контактного лица клиента.");
                Console.WriteLine("4 - Вывод клиента с наибольшим количеством заказов за указанный год и месяц");
                string command = Console.ReadLine();

                switch (command)
                {
                    case "2":
                        PrintClient(products, clients, bids);
                        break;
                    case "3":
                        EditClient(clients, filePath);
                        break;
                    case "4":
                        GetGoldClient(clients, bids);
                        break;
                }
            }
        }



        private static void ExtractProductsClientsBids(string filePath, out List<Product> products, out List<Client> clients, out List<Bid> bids)
        {
            products = new List<Product>();
            clients = new List<Client>();
            bids = new List<Bid>();
            var workbook = new XLWorkbook(filePath);

            // Товары
            var worksheet = workbook.Worksheet(1);
            // Получение всех строк из файла
            var rows = worksheet.RangeUsed().RowsUsed().Skip(1); // Skip() - Пропуск заголовка таблицы
            // Заполнение списка всех Товаров (путем получения строк из файла с Товарами)
            foreach (var row in rows)
            {
                Product product = new Product();
                product.Id = row.Cell(1).Value.ToString();
                product.Name = row.Cell(2).Value.ToString();
                product.Unit = row.Cell(3).Value.ToString();
                product.Price = row.Cell(4).Value.ToString();
                products.Add(product);
            }

            // Клиенты
            worksheet = workbook.Worksheet(2);
            // Получение всех строк из файла
            rows = worksheet.RangeUsed().RowsUsed().Skip(1);
            // Заполнение списка всех Клиентов
            foreach (var row in rows)
            {
                Client client = new Client();
                client.Id = row.Cell(1).Value.ToString();
                client.Name = row.Cell(2).Value.ToString();
                client.Address = row.Cell(3).Value.ToString();
                client.Contact = row.Cell(4).Value.ToString();
                clients.Add(client);
            }

            // Заявки
            worksheet = workbook.Worksheet(3);
            // Получение всех строк из файла
            rows = worksheet.RangeUsed().RowsUsed().Skip(1);
            // Заполнение списка всех Заявок
            foreach (var row in rows)
            {
                Bid application = new Bid();
                application.Id = row.Cell(1).Value.ToString();
                application.ProductId = row.Cell(2).Value.ToString();
                application.ClientId = row.Cell(3).Value.ToString();
                application.Number = row.Cell(4).Value.ToString();
                application.Quantity = row.Cell(5).Value.ToString();
                application.Date = row.Cell(6).GetDateTime();
                bids.Add(application);
            }
        }



        private static void PrintClient(List<Product> products, List<Client> clients, List<Bid> bids)
        {
            Console.WriteLine();
            Console.WriteLine("Введите наименование товара:");
            string productName = Console.ReadLine();
            Console.WriteLine();

            var existsProducts = products.Where(product => product.Name == productName).ToList();
            if (existsProducts.Count == 0)
            {
                Console.WriteLine($"Товар '{productName}' отсутствует в списке!");
                Console.WriteLine("\r\n \r\n");
                return;
            }

            var ProductsClientsBids = from product in existsProducts
                                      join bid in bids on product.Id equals bid.ProductId
                                      join client in clients on bid.ClientId equals client.Id
                                      select new
                                      {
                                          client.Name,
                                          client.Address,
                                          client.Contact,
                                          bid.Quantity,
                                          product.Unit,
                                          product.Price,
                                          bid.Date
                                      };

            Console.WriteLine("Информация о клиентах, заказавших товар: " + productName);
            Console.WriteLine("----------------------------------------");
            foreach (var item in ProductsClientsBids.ToList())
            {
                Console.WriteLine("- Наименование организации: " + item.Name);
                Console.WriteLine("- Адрес: " + item.Address);
                Console.WriteLine("- Контактное лицо (ФИО): " + item.Contact);
                Console.WriteLine("- Количество товара: " + item.Quantity);
                Console.WriteLine("- Единица измерения товара: " + item.Unit);
                Console.WriteLine("- Цена товара за единицу: " + item.Price);
                Console.WriteLine("- Дата заказа: " + item.Date);
                Console.WriteLine();
            }
            Console.WriteLine();
            Console.WriteLine("Нажмите любую клавишу, чтобы вернуться на главную...");
            Console.ReadKey();
            Console.WriteLine("\r\n \r\n");
        }



        private static void EditClient(List<Client> clients, string filePath)
        {
            Console.WriteLine();
            Console.WriteLine("Введите код клиента:");
            string clientId = Console.ReadLine();
            Console.WriteLine();

            var workbook = new XLWorkbook(filePath);
            // Клиенты
            var worksheet = workbook.Worksheet(2);
            // Получение всех строк из файла
            var rows = worksheet.RangeUsed().RowsUsed().Skip(1);

            // Перебор строк из файла
            foreach (var row in rows)
            {
                if (row.Cell(1).Value.ToString() == clientId)
                {
                    Console.WriteLine("Изменение контактного лица клиента:");
                    Console.WriteLine("-----------------------------------");
                    Console.WriteLine("Код клиента: " + clientId);
                    Console.WriteLine("Название организации: " + row.Cell(2).Value.ToString());
                    Console.WriteLine("Адрес: " + row.Cell(3).Value.ToString());
                    Console.WriteLine("Контактное лицо (ФИО): " + row.Cell(4).Value.ToString());
                    Console.WriteLine();

                    Console.WriteLine("Введите новое название организации:");
                    string companyName = Console.ReadLine();
                    if (companyName == "")
                    {
                        Console.WriteLine("Вы не ввели новое название организации!");
                        Console.WriteLine("\r\n \r\n");
                        return;
                    }

                    Console.WriteLine("Введите новое контактное лицо (ФИО):");
                    string address = Console.ReadLine();
                    if (address == "")
                    {
                        Console.WriteLine("Вы не ввели новое контактное лицо!");
                        Console.WriteLine("\r\n \r\n");
                        return;
                    }

                    // Изменение названия организации 
                    row.Cell(2).Value = companyName;
                    // Изменение контактного лица (ФИО)
                    row.Cell(4).Value = address;
                    workbook.Save();

                    Console.WriteLine();
                    Console.WriteLine("Название организации и Контактное лицо были успешно изменены!");
                    Console.WriteLine("Новое название организации: " + row.Cell(2).Value.ToString());
                    Console.WriteLine("Новое контактное лицо (ФИО): " + row.Cell(4).Value.ToString());
                    Console.WriteLine();
                    Console.WriteLine("Нажмите любую клавишу, чтобы вернуться на главную...");
                    Console.ReadKey();
                    Console.WriteLine("\r\n \r\n");
                    return;
                }
            }
            Console.WriteLine("Клиент с кодом '" + clientId + "' не найден!");
            Console.WriteLine();
            Console.WriteLine("Нажмите любую клавишу, чтобы вернуться на главную...");
            Console.ReadKey();
            Console.WriteLine("\r\n \r\n");
        }



        private static void GetGoldClient(List<Client> clients, List<Bid> bids)
        {
            Console.WriteLine("Введите год:");
            bool canParseYear = int.TryParse(Console.ReadLine(), out int year);
            if (canParseYear != true | year < 1970 | year > 2200)
            {
                Console.WriteLine("Введено неверное значение года!");
                Console.WriteLine("\r\n \r\n");
                return;
            }

            Console.WriteLine("Введите месяц:");
            bool canParseMonth = int.TryParse(Console.ReadLine(), out int month);
            if (canParseMonth != true | month < 1 | month > 12)
            {
                Console.WriteLine("Введено неверное значение месяца!");
                Console.WriteLine("\r\n \r\n");
                return;
            }

            DateTime startDate = new DateTime(year, month, 1);
            DateTime endDate = startDate.AddMonths(1).AddDays(-1);

            var BidsClients = from bid in bids
                              where (bid.Date >= startDate) && (bid.Date <= endDate)
                              join client in clients on bid.ClientId equals client.Id
                              select new
                              {
                                  client.Name,
                                  client.Contact,
                                  bid.Id,
                                  bid.Quantity,
                                  bid.Date,
                              } into bidsClients
                              group bidsClients by bidsClients.Name into groupBidsClients
                              orderby groupBidsClients.Key.Count() descending
                              select groupBidsClients;
            var firstBidClient = BidsClients.FirstOrDefault().ToList();

            Console.WriteLine();
            Console.WriteLine("Клиент с наибольшим количеством заказов за период " + startDate.ToString("d") + " - " + endDate.ToString("d") + ":");

            string clientName = firstBidClient.Select(client => client.Name).FirstOrDefault();
            Console.WriteLine("- Наименование организации клиента: " + clientName);

            string clientContact = firstBidClient.Select(client => client.Contact).FirstOrDefault();
            Console.WriteLine("- Контактное лицо клиента: " + clientContact);

            Console.WriteLine("- Заявки:");
            foreach (var item in firstBidClient)
            {
                Console.Write("     Код заявки: " + item.Id);
                Console.Write(", Количество товаров: " + item.Quantity);
                Console.Write(", Дата заявки: " + item.Date.ToString("d"));
                Console.WriteLine();
            }
            Console.WriteLine();
            Console.WriteLine("Нажмите любую клавишу, чтобы вернуться на главную...");
            Console.ReadKey();
            Console.WriteLine("\r\n \r\n");
        }




    }
}