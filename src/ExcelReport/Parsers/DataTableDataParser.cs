using System;
using System.Text.RegularExpressions;

namespace ExcelReport.Parsers
{
    public class DataTableDataParser : RegexParser
    {
        private static readonly Regex regex = new Regex(@"(?<=\$\<)([\w]*)(?=\>\[data\])");

        public override Regex Regex => regex;
    }
}
