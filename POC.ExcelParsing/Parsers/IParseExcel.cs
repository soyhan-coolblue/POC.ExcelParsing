using System.Collections.Generic;

using POC.ExcelParsing.Dto;

namespace POC.ExcelParsing.Parsers
{
    public interface IParseExcel
    {
        IReadOnlyCollection<ProductPrice> Parse(byte[] excelFile);
    }
}
