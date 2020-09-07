using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

using OfficeOpenXml;
using OfficeOpenXml.Table;

using POC.ExcelParsing.Dto;


namespace POC.ExcelParsing.Parsers
{
    public class OfficeOpenXmlParser : IParseExcel
    {
        private const string _sheetName = "Universal template";

        public IReadOnlyCollection<ProductPrice> Parse(byte[] excelFile)
        {
            var excelTable = GetExcelSheet(excelFile);

            return GetProductPrices(excelTable);
        }

        private ExcelWorksheet GetExcelSheet(byte[] excelFile)
        {
            var package = new ExcelPackage(new MemoryStream(excelFile));
            ExcelWorksheet sheet = package.Workbook.Worksheets[_sheetName];

            return sheet;
        }

        private IReadOnlyCollection<ProductPrice> GetProductPrices(ExcelWorksheet excelWorksheet)
        {
            var productPrices = new List<ProductPrice>();

            // Starts from 4th row.
            int index = 4;

            do
            {
                var productPrice = new ProductPrice()
                {
                    ProductId = excelWorksheet.Cells[index, ColumnIndexes.ProductId].GetValue<string>(),
                    Price = excelWorksheet.Cells[index, ColumnIndexes.Price].GetValue<decimal>(),
                    EanCode = excelWorksheet.Cells[index, ColumnIndexes.EanCode].GetValue<string>(),
                    ManufacturerCode = excelWorksheet.Cells[index, ColumnIndexes.ManufacturerCode].GetValue<string>(),
                    StartDate = excelWorksheet.Cells[index, ColumnIndexes.StartDate].GetValue<DateTime>(),
                    EndDate = excelWorksheet.Cells[index, ColumnIndexes.EndDate].GetValue<DateTime>(),
                };

                productPrices.Add(productPrice);

                index++;

            } while (excelWorksheet.Cells[index, ColumnIndexes.ProductId].Value != null);

            return productPrices;
        }
    }

    internal class ColumnIndexes
    {
        public static int ProductId = 1;
        public static int EanCode = 2;
        public static int ManufacturerCode = 3;
        public static int Price = 4;
        public static int StartDate = 5;
        public static int EndDate = 6;
        public static int PriceChangeType = 7;
    }
}
