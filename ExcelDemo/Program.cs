using ClosedXML.Excel;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;

namespace ExcelDemo
{
    public class Program
    {
        static async Task Main(string[] args)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            var file = new FileInfo(@"C:\Demos\Youtube.xlsx");

            var people = GetDataSetUp();

            await SaveFile(people, file);

            // new Program().Export<PersonModel>(people, @"C:\ExcelDemo\Exportdata.xlsx", "people");
            List<PersonModel> peopleFromExcel = await LoadFromExcelFile(file);
                 foreach (var p in peopleFromExcel) Console.WriteLine($"{p.Id} {p.FirstName} {p.LastName}");
          
        }

        //public bool Export<T>(List<T> list, string file, string sheetName)
        //{
        //    bool exported = false;

        //    using(IXLWorkbook workbook = new XLWorkbook(file))
        //    {
        //        workbook.AddWorksheet(sheetName).FirstCell().InsertTable<T>(list,false);

        //        workbook.SaveAs(file);

        //        exported = true;
        //    }

        //    return exported;
        //}

        private static async Task<List<PersonModel>> LoadFromExcelFile(FileInfo file)
        {
            List<PersonModel> output = new();
            using var package = new ExcelPackage(file);

            await package.LoadAsync(file);

            var ws = package.Workbook.Worksheets[0];

            int row = 2;
            int col = 1;

            while (string.IsNullOrEmpty(ws.Cells[row, col].Value?.ToString()) == false)
            {
                PersonModel person = new PersonModel();

                person.Id = int.Parse(ws.Cells[row, col].Value.ToString());
                person.FirstName = ws.Cells[row, col + 1].Value.ToString();
                person.LastName = ws.Cells[row, col + 2].Value.ToString();

                output.Add(person);

                row++;
            }

            return output;
        }

        private static async Task SaveFile(List<PersonModel> people, FileInfo file)
        {
            DeleteIfExist(file);

          using  var package = new ExcelPackage(file);

            var ws = package.Workbook.Worksheets.Add("MainReport");

            var range = ws.Cells["A1"].LoadFromCollection(people, true);
            range.AutoFitColumns();

            ws.Column(1).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

           await package.SaveAsync();
        }

        private static void DeleteIfExist(FileInfo file)
        {
           if(file.Exists) file.Delete();
        }

        static List<PersonModel> GetDataSetUp()
        {
            List<PersonModel> result = new()
            {
                new(){Id = 1, FirstName = "Stanley", LastName = "Ukwuze"},
                new(){Id = 2, FirstName = "Chris", LastName = "Jame"},
                new(){Id = 3, FirstName = "Faith", LastName = "Favour"}
            };

            return result;
        }
    }
}
