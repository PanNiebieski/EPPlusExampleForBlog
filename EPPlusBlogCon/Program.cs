using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPlusBlogCon
{
    class Program
    {
        static void Main(string[] args)
        {
            DirectoryInfo outputDir = new DirectoryInfo(@"d:\Samples\EPPlus\");

            if (!outputDir.Exists)
                outputDir.Create();

            CreateExcelFile<Game>.ExportToExcel(CreateCollection(), outputDir, "Games");

            Console.WriteLine("File Created");
            Console.ReadKey();

        }

        public static List<Game> CreateCollection()
        {
            List<Game> list = new List<Game>();

            list.Add(new Game() { Name = "Mortal Kombat",Price = 120,SellAmount=120853});
            list.Add(new Game() { Name = "Asura Wrath", Price = 90, SellAmount = 90045 });
            list.Add(new Game() { Name = "X-COM", Price = 210, SellAmount = 117651 });
            list.Add(new Game() { Name = "DMC", Price = 210, SellAmount = 81123 });

            return list;
        }
    }

    public static class CreateExcelFile<T>
    {
        public static void ExportToExcel(IEnumerable<T> employees, DirectoryInfo outputDir, string fileName)
        {
            fileName = FixFileNameExcel(fileName);

            FileInfo f = new FileInfo(outputDir.FullName + @"\" + fileName);
            DeleteFileIfExist(f);
            using (var excelFile = new ExcelPackage(f))
            {
                var worksheet = excelFile.Workbook.Worksheets.Add("Games");
                worksheet.Cells["A1"].LoadFromCollection(Collection: employees, PrintHeaders: true);
                worksheet.Cells.AutoFitColumns(0);
                excelFile.Save();
            }
        }

        public static FileInfo DeleteFileIfExist(FileInfo newFile)
        {
            if (newFile.Exists)
            {
                newFile.Delete();  // ensures we create a new workbook
                newFile = new FileInfo(newFile.FullName);
            }
            return newFile;
        }

        public static string FixFileNameExcel(string fileName)
        {
            if (Path.HasExtension(fileName))
            {
                string ext = Path.GetExtension(fileName);

                if (ext != ".xlsx")
                {
                    fileName = Path.GetFileNameWithoutExtension(fileName) + ".xlsx";
                }
            }
            else
            {
                fileName += ".xlsx";
            }

            return fileName;
        }
    }




    public class Game
    {
        public string Name { get; set; }
        public decimal Price { get; set; }
        public int SellAmount { get; set; }
    }
}
