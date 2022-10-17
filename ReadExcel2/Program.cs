using ExcelDataReader;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static System.Console;

namespace ReadExcel2
{
    class Program
    {
        static void Main(string[] args)
        {
            var localInDirectory = "C:\\\\Users\\\\weichienyap\\\\Desktop\\\\";
            var localFile = "Sold&Balance-T&L_PREMIUM.xls"; //"TestExcelData.xlsx";
            var finalFilePath = localInDirectory + localFile;

            var fileExtension = Path.GetExtension(finalFilePath).ToUpper();
            var validFilePath = new List<string>() { ".XLS", ".XLSX" };

            if (!validFilePath.Contains(fileExtension))
            {
                WriteLine("Invalid Extension");
                Read();
                return;
            }

            using (var stream = File.Open(finalFilePath, FileMode.Open, FileAccess.Read))
            {
                using (var reader = ExcelReaderFactory.CreateReader(stream))
                {
                    var result = reader.AsDataSet();

                    if (result.Tables == null || result.Tables.Count == 0)
                    {
                        WriteLine("Empty excel");
                        Read();
                        return;
                    }

                    var dtData = result.Tables[0];
                    var rowCount = dtData.Rows.Count - 2; //Minus 2 due to last two row in excel not using by the system
                    var columnCount = dtData.Columns.Count;
                    var rows = dtData.Select();

                    var columnIndex = 11; //First column for center code
                    var columnIndexList = new List<ExcelColumnIndex>();
                    var excelDataList = new List<ExcelDataModel>();

                    // Calculate all the center code & sales quantity column index
                    while (columnCount > columnIndex)
                    {
                        columnIndexList.Add(new ExcelColumnIndex() { CenterCodeColumnIndex = columnIndex, SalesQuantityColumnIndex = columnIndex + 1 });
                        columnIndex += 3;
                    }

                    // Start extract data from the excel
                    var itemStartRowIndex = 4; // Item code will start at row number 5 in Excel
                    var itemCodeColumnIndex = 1;
                    var itemNameColumnIndex = 2;
                    var centerCodeRowIndex = 0;
                    for (int i = itemStartRowIndex; i < rowCount; i++)
                    {
                        foreach (var indexItem in columnIndexList)
                        {
                            excelDataList.Add(new ExcelDataModel()
                            {
                                CenterCode = rows[centerCodeRowIndex][indexItem.CenterCodeColumnIndex].ToString(),
                                ItemCode = rows[i][itemCodeColumnIndex].ToString(),
                                SalesQty = Int32.Parse(rows[i][indexItem.SalesQuantityColumnIndex].ToString()),
                                ItemName = rows[i][itemNameColumnIndex].ToString()
                            });
                        }
                    }

                    //foreach (var item in excelDataList)
                    //{
                    //    WriteLine($"{item.CenterCode}, {item.ItemCode}, {item.ItemName}, {item.SalesQty}");
                    //}

                    //WriteLine(excelDataList.Count);           
                }
            }

            Read();
        }

        public class ExcelColumnIndex
        {
            public int CenterCodeColumnIndex { get; set; }
            public int SalesQuantityColumnIndex { get; set; }
        }

        public class ExcelDataModel
        {
            public string CenterCode { get; set; }
            public string ItemCode { get; set; }
            public string ItemName { get; set; }
            public int SalesQty { get; set; }
        }
    }
}
