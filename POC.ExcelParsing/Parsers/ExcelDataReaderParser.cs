using System;
using System.Collections.Generic;
using System.Data;
using System.IO;

using ExcelDataReader;

using POC.ExcelParsing.Dto;

namespace POC.ExcelParsing.Parsers
{
    public class ExcelDataReaderParser : IParseExcel
    {
        private const string _sheetName = "Universal template";

        public IReadOnlyCollection<ProductPrice> Parse(byte[] excelFile)
        {
            //return ParseUsingDataset(excelFile);

            return ParseUsingReader(excelFile);
        }

        private IReadOnlyCollection<ProductPrice> ParseUsingDataset(byte[] excelFile)
        {
            using var stream = new MemoryStream(excelFile);

            using var reader = ExcelReaderFactory.CreateReader(stream);

            var dataSet = reader.AsDataSet();

            var dataTable = GetDataTable(dataSet);

            return ParseDataTable(dataTable);
        }

        private DataTable GetDataTable(DataSet dataSet)
        {
            return dataSet.Tables[_sheetName];
        }

        private IReadOnlyCollection<ProductPrice> ParseDataTable(DataTable dataTable)
        {
            var productPrices = new List<ProductPrice>();
            int dataStartsAt = 3;

            for (int index = dataStartsAt; index < dataTable.Rows.Count; index++)
            {
                DataRow row = dataTable.Rows[index];

                if(row[ColumnConstants.ProductId] == null)
                    continue;
                
                var productPrice = new ProductPrice()
                {
                    ProductId = row[ColumnConstants.ProductId].ToString(),
                    EanCode = row[ColumnConstants.EanCode].ToString(),
                    ManufacturerCode = row[ColumnConstants.ManufacturerCode].ToString(),
                    Price = decimal.Parse(row[ColumnConstants.Price].ToString()),
                    StartDate = DateTime.Parse(row[ColumnConstants.StartDate].ToString()),
                    EndDate = DateTime.Parse(row[ColumnConstants.EndDate].ToString())
                };

                productPrices.Add(productPrice);
            }

            return productPrices;
        }

        private IReadOnlyCollection<ProductPrice> ParseUsingReader(byte[] excelFile)
        {
            var productPrices = new List<ProductPrice>();

            using var stream = new MemoryStream(excelFile);

            using var reader = ExcelReaderFactory.CreateReader(stream);

            do
            {
                while (reader.Read())
                {
                    if (reader.Depth < 3)
                        continue;

                    if (reader.GetValue(ColumnConstants.ProductId) == null) 
                        continue;

                    var productPrice = new ProductPrice()
                    {
                        ProductId = reader.GetValue(ColumnConstants.ProductId).ToString(),
                        ManufacturerCode = reader.GetValue(ColumnConstants.ManufacturerCode).ToString(),
                        EanCode = reader.GetValue(ColumnConstants.EanCode).ToString(),
                        Price = decimal.Parse(reader.GetValue(ColumnConstants.Price).ToString()),
                        StartDate = reader.GetDateTime(ColumnConstants.StartDate),
                        EndDate = reader.GetDateTime(ColumnConstants.EndDate)
                    };

                    productPrices.Add(productPrice);
                }
            } while (reader.NextResult() && reader.Name == _sheetName);

            return productPrices;
        }
    }

    public class ColumnConstants
    {
        public static int ProductId = 0;
        public static int EanCode = 1;
        public static int ManufacturerCode = 2;
        public static int Price = 3;
        public static int StartDate = 4;
        public static int EndDate = 5;
        public static int PriceChangeType = 6;
    }
}