using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeOpenXml;

namespace ImageNameToExcel
{
  public static class FileToExcel
  {
    public static async Task SaveExcelFile(List<PersonModel> person, FileInfo file)
    {
     
      DeleteIfExists(file);
      Console.WriteLine("================> Writing Names Excel File...<=================");
      ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
      using var package = new ExcelPackage(file);

      var ws = package.Workbook.Worksheets.Add("MainData");

      var range = ws.Cells["A2"].LoadFromCollection(person, true);
      range.AutoFitColumns();

      //Formats the header
      ws.Cells["A1"].Value = "Reference Numbers";
      await package.SaveAsync();
      Console.WriteLine("================> Writing Completed...<=================");
    }

    private static void DeleteIfExists(FileInfo file)
    {
      Console.WriteLine("================> Checking and Deleting Duplicate File name...<=================");
      if (file.Exists)
      {
        Console.WriteLine("================> File Found and Deleted...<=================");
        file.Delete();
      }

      Console.WriteLine("================> Delete Action Completed...<=================");
    }
  }
}
