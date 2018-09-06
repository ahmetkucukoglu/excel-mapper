namespace ExcelMappingApp.Library
{
    using System;
    using System.Text.RegularExpressions;

    public class ExcelUtilities
    {
        private const string COLUMN_REGEX_PATTERN = @"([A-Za-z]+)(\d+)";

        public static Tuple<string, int> GetPartsOfColumn(string column)
        {
            EnsureColumnValid(column);

            var regex = new Regex(COLUMN_REGEX_PATTERN);
            var match = regex.Match(column);

            var columnName = match.Groups[1].Value;
            var columnIndex = match.Groups[2].Value;

            return new Tuple<string, int>(match.Groups[1].Value, int.Parse(match.Groups[2].Value));
        }

        public static void EnsureColumnValid(string column)
        {
            var regex = new Regex(COLUMN_REGEX_PATTERN);
            var regexMatch = regex.Match(column);

            if (!regexMatch.Success)
                throw new NotSupportedColumnFormatException("NotSupportedColumnFormat");
        }
    }
}
