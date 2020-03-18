using System;
using System.Text.RegularExpressions;

namespace ExcelReport.Parsers
{
    public sealed class DataTableHeaderParser:RegexParser
    {
        private static readonly Regex regex = new Regex(@"(?<=\$\<)([\w]*)(?=\>\[header\])");

        public override Regex Regex => regex;
    }
}
