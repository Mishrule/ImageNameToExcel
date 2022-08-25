using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ImageNameToExcel
{
  public static class CopyFromFile
  {
    private static readonly List<PersonModel> _personNames = new();
    public static List<PersonModel> CopyFromFiles()
    {
      string sourceFolderName = @"C:\Users\mishr\Desktop\Testing\sourceFolder";
      //string destinationFolder = @"C:\Users\mishr\Desktop\Testing\destinationFolder";

      string[] getAllFiles = Directory.GetFiles(sourceFolderName);
      Console.WriteLine("================> Getting fileNames from files...<=================");
      foreach (var file in getAllFiles)
      {
        var fileName = Path.GetFileNameWithoutExtension(file);
        //var extension =   Path.GetExtension(file);

        _personNames.Add(new()
        {
          Name = fileName
        });
      }

      Console.WriteLine("================> Name Completed... <=================");
      return _personNames;

    }
  }
}
