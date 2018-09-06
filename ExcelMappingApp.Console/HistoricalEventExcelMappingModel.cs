namespace ExcelMappingApp.Console
{
    using ExcelMappingApp.Library;
    using System;

    public class HistoricalEventExcelMappingModel : IExcelMappingModel
    {
        public string Row { get; set; }
        public string Name  { get; set; }
        public decimal Booty { get; set; }
        public DateTime Date { get; set; }
    }
}
