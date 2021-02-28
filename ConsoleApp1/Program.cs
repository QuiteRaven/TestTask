using AutoMapper;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;


namespace TestTask
{
    class Program
    {
        static async Task Main(string[] args)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            var file = new FileInfo("Data.xlsx");
            List<Person> data = GetData();

            await SaveFile(data, file);

            List<Person> persons = await ReadFile(file);
            List<PersonDb> personsDb = await MapEntitiesToDb(persons);
            DBContext dBContext = new DBContext();
            await dBContext.persons.AddRangeAsync(personsDb);
            await dBContext.SaveChangesAsync();

        }

        private static async Task SaveFile(List<Person> data, FileInfo file)
        {
            DeleteExistents(file);

            using var package = new ExcelPackage(file);

            var ws = package.Workbook.Worksheets.Add("Main");

            ws.Cells["A1"].Value = "First Name";
            ws.Cells["B1"].Value = "Second Name";
            ws.Cells["C1"].Value = "Middle Name";
            ws.Cells["D1"].Value = "Phone Number";
            ws.Cells["E1"].Value = "Address";

            var range = ws.Cells["A2"].LoadFromCollection(data,false);
            range.AutoFitColumns();

            await package.SaveAsync();
        }

        private static async Task<List<Person>> ReadFile(FileInfo file)
        {
            List<Person> valuesList = new ();
            using var package = new ExcelPackage(file);

            var ws = package.Workbook.Worksheets[0];
            for (int i = ws.Dimension.Start.Row + 1; i <= ws.Dimension.End.Row; i++)
            {
                var value = new Person();

                value.FirstName = ws.Cells[i, 1].Value.ToString();
                value.SecondName = ws.Cells[i, 2].Value.ToString();
                value.MiddleName = ws.Cells[i, 3].Value.ToString();
                value.PhoneNumber = ws.Cells[i, 4].Value.ToString();
                value.Address = ws.Cells[i, 5].Value.ToString();
                
                valuesList.Add(value);
            }

            return await Task.FromResult(valuesList);
        }

        private static async Task<List<PersonDb>> MapEntitiesToDb(List<Person> persons) 
        {
            var config = new MapperConfiguration(config => config.CreateMap<Person, PersonDb>());
            var mapper = new Mapper(config);
            
            List<PersonDb> personsDb = new();

            for(int i = 0; i<persons.Count; i++) 
            {
                var persondb = mapper.Map<PersonDb>(persons[i]);
                personsDb.Add(persondb);
            }

            return await Task.FromResult(personsDb);
        }

        private static void DeleteExistents(FileInfo file)
        {
            if (file.Exists)
                file.Delete();
        }

        private static List<Person> GetData()
        {
            Random random = new Random();
            List<Person> output = new();

            for (int i = 0; i < 200000; i++)
            {
                var value = Guid.NewGuid().ToString();
                output.Add(new Person
                {
                    FirstName = "Oleg"+i,
                    SecondName = "Ivanov"+i,
                    MiddleName = "Maratovich"+i,
                    PhoneNumber = random.Next(45648494).ToString(),
                    Address = "Tokyo"+i
                });
            }
     
            return output;
        }

    }
}
