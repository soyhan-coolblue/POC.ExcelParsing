using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;

using POC.ExcelParsing.Dto;
using POC.ExcelParsing.Parsers;

namespace POC.ExcelParsing.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class UploadController : ControllerBase
    {
        public UploadController()
        {
            
        }

        [HttpPost]
        [Route("ForOfficeOpenXml")]
        public async Task<IActionResult> ForOfficeOpenXml(IFormFile excelFile, 
            [FromServices] OfficeOpenXmlParser officeOpenXmlParser)
        {
            var result = new ProductPricesWithTime();

            var stopWatch = Stopwatch.StartNew();

            result.ProductPrices = officeOpenXmlParser.Parse(await excelFile.GetBytes());
            stopWatch.Stop();
            result.ParsingTimeInMillis = stopWatch.Elapsed.TotalMilliseconds;

            return Ok(result);
        }

        [HttpPost]
        [Route("ForExcelDataReader")]
        public async Task<IActionResult> ForExcelDataReader(IFormFile excelFile, 
            [FromServices] ExcelDataReaderParser excelReader)
        {
            var result = new ProductPricesWithTime();

            var stopWatch = Stopwatch.StartNew();

            result.ProductPrices = excelReader.Parse(await excelFile.GetBytes());
            stopWatch.Stop();
            result.ParsingTimeInMillis = stopWatch.Elapsed.TotalMilliseconds;
            
            return Ok(result);
        }
    }

    public static class FormFileExtensions
    {
        public static async Task<byte[]> GetBytes(this IFormFile formFile)
        {
            using (var memoryStream = new MemoryStream())
            {
                await formFile.CopyToAsync(memoryStream);
                return memoryStream.ToArray();
            }
        }
    }

    public class ProductPricesWithTime{
        public IReadOnlyCollection<ProductPrice> ProductPrices{ get; set; }
        public double ParsingTimeInMillis { get; set; }
    }
}
