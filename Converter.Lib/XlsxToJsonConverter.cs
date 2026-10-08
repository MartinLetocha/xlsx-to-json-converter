using System.Text.Json;
using ClosedXML.Excel;

namespace Converter.Lib;

public class XlsxToJsonConverter
{
    public string Convert(FileInfo spreadsheet)
    {
        List<Client> clients = new List<Client>();
        var workbook = new XLWorkbook(spreadsheet.FullName).Worksheet(1);
        int rowNumber = 2;
        int cellNumber = 4;
        var titleRow = workbook.Row(1);
        
        while (true)
        {
            Client client;
            var row = workbook.Row(rowNumber);
            var clientName = row.Cell(1).Value.ToString();
            
            if (String.IsNullOrEmpty(clientName))
            {
                break;
            }

            var clientFromList = clients.FirstOrDefault(x => x.Name == clientName);
            
            if (clientFromList == null)
            {
                client = new Client() { Name = clientName, Orders = new List<Order>() };
                clients.Add(client);
            }
            else
            {
                client = clientFromList;
            }

            client.ID = row.Cell(2).Value.ToString();

            Order order = new Order
            {
                Items = new List<Item>(),
                OrderName = row.Cell(3).Value.ToString()
            };

            while (true)
            {
                Item item = new Item();
                var cell = row.Cell(cellNumber);
                string value = cell.Value.ToString();
                string period = titleRow.Cell(cellNumber).Value.ToString();

                if (String.IsNullOrEmpty(period))
                {
                    break;
                }

                item.Period = DateTime.Parse(period).ToShortDateString();
                item.Amount = String.IsNullOrEmpty(value) ? 0 : int.Parse(value);

                order.Items.Add(item);
                cellNumber++;
            }

            client.Orders.Add(order);
            cellNumber = 4;
            rowNumber++;
        }

        return JsonSerializer.Serialize(clients, new JsonSerializerOptions { WriteIndented = true });
    }
}

public class Client
{
    public string Name { get; set; }
    public string ID { get; set; }
    public List<Order> Orders { get; set; }
}

public class Order
{
    public string OrderName { get; set; }
    public List<Item> Items { get; set; }
}

public class Item
{
    public string Period { get; set; }
    public int Amount { get; set; }
}