using Newtonsoft.Json;
using System;

namespace File.Poc
{
    class Program
    {
        static void Main(string[] args)
        {
            var path = @"C:\Tmp\users.xlsx";
            var fileService = new FileService();

            var list = fileService.ListToExcelSax<FileDto>(path);
            //Console.Write($"{JsonConvert.SerializeObject(list)}");

            //var list = fileService.ListToExcelDOM<FileDto>(path);
            //Console.Write($"{JsonConvert.SerializeObject(list)}");

            //var list = fileService.ImportExcel<FileDto>(path);
            //Console.Write($"{JsonConvert.SerializeObject(list)}");

            //var dt = fileService.ExtractExcelSheetValuesToDataTable(path);
            //Console.Write($"{JsonConvert.SerializeObject(dt)}");

            //var dt = fileService.ExtractExcelSAXToDataTable(path);
            //Console.Write($"{JsonConvert.SerializeObject(dt)}");

            if (list.Count > 0) Console.Write($"Lista de {list.Count} gerada com sucesso!");
        }
    }
}
