namespace ExcelMappingApp.Library
{
    using System;

    public class NotMatchHeaderException : Exception
    {
        public NotMatchHeaderException(string message) : base(message) { }
    }
}
