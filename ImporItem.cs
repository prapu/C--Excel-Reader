using ExcelDataReader;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ReadExcel
{
    public class ImportItem
    {
        public int ItemNumber { get; set; }
        public string Description { get; set; }
        public string Unit { get; set; }
        public decimal Price { get; set; }
        public int OrderQuantity { get; set; }
        public int QuantityReceived { get; set; }
        public DateTime? OrderDate { get; set; }
        public DateTime? ReceivedDate { get; set; }
        public ImportItem() { }

        public static List<ImportItem> ImportItems()
        {
            //list of items
            List<ImportItem> lst = new List<ImportItem>();
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

            //Make sure change this path 
            var p = Path.GetFullPath(@"C:\Praveen\vs2019\ReadExcel\");

            //create a file stream 
            using (var stream = File.Open(p + "ImportItems.xlsx", FileMode.Open, FileAccess.Read))
            {
                //read the file stream
                using (var reader = ExcelReaderFactory.CreateReader(stream))
                {
                    do
                    {
                        while (reader.Read()) //Each ROW
                        {
                            //ignoring the heading
                            if (reader.Depth > 0)
                            {
                                var bl = new ImportItem();
                                for (int column = 0; column < reader.FieldCount; column++)
                                {
                                    switch (column)
                                    {
                                        case 0://1 ITem Number
                                            bl.ItemNumber =int.Parse(reader.GetValue(column).ToString());
                                            break;
                                        case 1: //Description
                                            bl.Description =reader.GetValue(column).ToString();
                                            break;
                                        case 2: //Unit of measure
                                            bl.Unit = reader.GetValue(column).ToString();
                                            break;
                                        case 3: //each Price
                                            bl.Price = decimal.Parse(reader.GetValue(column).ToString());
                                            break;
                                        case 4: //Order quantity
                                            bl.OrderQuantity = int.Parse(reader.GetValue(column).ToString());
                                            break;
                                        case 5: //Description
                                            bl.QuantityReceived = int.Parse(reader.GetValue(column).ToString());
                                            break;
                                        case 6: //Description
                                            bl.OrderDate = DateTime.Parse(reader.GetValue(column).ToString());
                                            break;
                                        case 7: //Description
                                            bl.ReceivedDate = DateTime.Parse(reader.GetValue(column).ToString());
                                            break;
                                    }
                                }
                                lst.Add(bl);
                            }

                        }
                    } while (reader.NextResult()); //Move to NEXT SHEET
                }
            }
            return lst;
        }
    }
}
