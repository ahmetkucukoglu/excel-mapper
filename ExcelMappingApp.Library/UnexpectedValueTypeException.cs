namespace ExcelMappingApp.Library
{
    using System;

    public class UnexpectedValueTypeException : Exception
    {
        public UnexpectedValueTypeException(string message) : base(message) { }
    }
}
