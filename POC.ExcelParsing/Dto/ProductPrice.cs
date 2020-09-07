using System;

namespace POC.ExcelParsing.Dto
{
    public class ProductPrice
    {
        public string ProductId { get; set; }
        public string EanCode { get; set; }
        public string ManufacturerCode { get; set; }
        public decimal Price { get; set; }
        public DateTime? StartDate { get; set; }
        public DateTime? EndDate { get; set; }
        public string PriceChangeType { get; set; }
    }
}
