using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace XLSXParser;

public class EPPlusRealization
{
    public void Process(Stream stream)
    {
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        try
        {
            var package = new ExcelPackage(stream);
            var sheet = package.Workbook.Worksheets[0];
            for (var i = 1; i < 33; ++i)
            {
                var data = sheet.Cells[3, i].Value.ToString();
                Console.WriteLine(data);
            }
        }

        catch (Exception ex)
        {
            var c = ex;
        }
    }
}
