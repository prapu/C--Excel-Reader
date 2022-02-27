using ExcelDataReader;
using System;
using System.Collections.Generic;
using System.IO;

namespace ReadExcel
{
    //this app uses ExcelDataReader Nuget package
    class Program
    {
        static void Main(string[] args)
        {

            foreach (var item in ImportItem.ImportItems())
            {
                Console.WriteLine($"Item Number: {item.ItemNumber}");
                Console.WriteLine($"Description: {item.Description}");
                Console.WriteLine($"Unit of Measure: {item.Unit}");
                Console.WriteLine($"Item Price: {item.Price}");
                Console.WriteLine($"Order Qty: {item.OrderQuantity}");
                Console.WriteLine($"Received Qty: {item.QuantityReceived}");
                Console.WriteLine($"Order Date: {item.OrderDate}");
                Console.WriteLine($"Received Date: {item.QuantityReceived}");
                Console.WriteLine(string.Empty);
            }

        }
    }
}
