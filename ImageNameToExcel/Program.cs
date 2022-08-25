using System;
using System.IO;
using System.Threading.Tasks;

namespace ImageNameToExcel
{
  public class Program
  {
    static async Task Main(string[] args)
    {
      Console.WriteLine("================> Initializing the model...<=================");
      var file = new FileInfo(@"C:\Users\mishr\Desktop\Testing\destinationFolder\Names.xlsx");
      await FileToExcel.SaveExcelFile(CopyFromFile.CopyFromFiles(), file);
    }
  }
}
