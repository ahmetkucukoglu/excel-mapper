namespace ExcelMappingApp.Library
{
    using System;

    public class NotSupportedColumnFormatException : Exception
    {
        public NotSupportedColumnFormatException(string message) : base(message) { }
    }
}
