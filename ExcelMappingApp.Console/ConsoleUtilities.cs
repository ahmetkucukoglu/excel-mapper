namespace ExcelMappingApp.Console
{
    using System;

    public class ConsoleUtilities
    {
        public const int TABLE_WIDTH = 95;
        public const string COLUMN_SEPERATE = "|";

        public static void PrintRow(params string[] columns)
        {
            int width = (TABLE_WIDTH - columns.Length) / columns.Length;
            string row = COLUMN_SEPERATE;

            foreach (string column in columns)
                row += AlignCenter(column, width) + COLUMN_SEPERATE;

            Console.WriteLine(row);
        }

        public static string AlignCenter(string text, int width)
        {
            text = text.Length > width ? text.Substring(0, width - 3) + "..." : text;

            if (string.IsNullOrEmpty(text))
                return new string(' ', width);
            else
                return text.PadRight(width - (width - text.Length) / 2).PadLeft(width);
        }
    }
}
