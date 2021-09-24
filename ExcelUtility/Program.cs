using ExcelUtility.Models;
using Newtonsoft.Json;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Threading.Tasks;

namespace ExcelUtility
{
    class Program
    {
        static async Task Main(string[] args)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            // Get the root folder, assuming that the system has already C:drive
            string root = @"C:\TestFiles";
            if (!Directory.Exists(root))
            {
                Directory.CreateDirectory(root);
            }
            // Concatenate the file 
            var filePath = root+@"\Test.xlsx";
            // The whole file path which will be used for saving the data into the excel sheet
            var file = new FileInfo(filePath);
            // Create a list of users
            var users = CreateData();
            // Finally generate excel sheet
            await GenerateExcelFile(users, file);
        }

        private static async Task GenerateExcelFile(List<UserModel> users, FileInfo file)
        {
            // Delete the existing file if already exists: The data will be overwrite
            DeleteIfExists(file);
            // Customized the properties according to the requirements
            using var excelFile = new ExcelPackage(file);
            var ws = excelFile.Workbook.Worksheets.Add("AllUsers");
            // Save the list of users into an excel sheet
            var range = ws.Cells["A1"].LoadFromCollection(users, true);
            range.AutoFitColumns();
            await excelFile.SaveAsync();
        }

        private static void DeleteIfExists(FileInfo file)
        {
           if(file.Exists)
            {
                file.Delete();
            }
        }

        private static List<UserModel> CreateData()
        {
            List<UserModel> users = new List<UserModel>()
            {
                new() { Id = 1, FirstName = "AKA", LastName = "Tariq" },
                new() { Id = 2, FirstName = "Shiraz", LastName = "Tariq" },
                new() { Id = 3, FirstName = "Akbar", LastName = "Ali" },
                new() { Id = 4, FirstName = "John", LastName = "Doe" },
            };
            return users;
        }
    
    }
}
