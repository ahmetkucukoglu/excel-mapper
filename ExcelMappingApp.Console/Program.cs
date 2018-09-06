namespace ExcelMappingApp.Console
{
    using ExcelMappingApp.Library;
    using System;
    using System.IO;

    class Program
    {
        static void Main(string[] args)
        {
            var bytes = File.ReadAllBytes("C:\\HistoricalEvents.xlsx");
            var excel = new MemoryStream(bytes);

            ExcelDataMapper<HistoricalEventExcelMappingModel> mapper = ExcelMapper();

            var historicalEvents = mapper.FromStream(excel).Map();
            
            ConsoleUtilities.PrintRow("SATIR", "OLAY", "TARİH", "GANİMET");

            foreach (var historicalEvent in historicalEvents)
            {
                ConsoleUtilities.PrintRow(historicalEvent.Row, historicalEvent.Name, historicalEvent.Date.ToShortDateString(), historicalEvent.Booty.ToString());
            }

            Console.ReadLine();
        }

        #region Configurations

        private static ExcelDataMapper<HistoricalEventExcelMappingModel> ExcelMapper() =>
           new ExcelDataMapper<HistoricalEventExcelMappingModel>()
               .Bind((x) => x.Name, "A1", "OLAY")
               .Bind((x) => x.Date, "B1", "TARİH", (x) =>
               {
                   return DateTime.Parse(x.ToString());
               })
               .Bind((x) => x.Booty, "C1", "GANİMET", (x) =>
               {
                   return decimal.Parse(x.ToString().Replace(".", ","));
               });

        #endregion
    }
}
