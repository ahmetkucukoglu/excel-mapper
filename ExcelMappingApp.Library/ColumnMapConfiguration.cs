namespace ExcelMappingApp.Library
{
    using System;

    public class ColumnMapConfiguration
    {
        public string PropertyName { get; set; }
        public string HeaderAddress { get; set; }
        public string HeaderTitle { get; set; }
        public Func<object, object> Converter { get; set; }
    }
}
