using System.Text.RegularExpressions;

namespace SnippetDesignerComponents
{
    public static class SnippetRegexPatterns
    {
        private const string PotentialReplacementStringFormat = @"^{0}$";
        private const string ReplacementStringFormat = @"((?<!{1}){1}{0}{1})|((?<={1}{0}{1}){1}{0}{1})";
        private const string ReplacementPart = @"(("".*"")|(\w+))";

        public static Regex BuildValidPotentialReplacementRegex()
        {
            var validPotentialReplacement = string.Format(PotentialReplacementStringFormat, ReplacementPart);
            return new Regex(validPotentialReplacement, RegexOptions.Compiled);
        }

        public static Regex BuildValidReplacementRegex(string delimiter)
        {
            var validReplacementString = BuildValidReplacementString(delimiter);
            return new Regex(validReplacementString, RegexOptions.Compiled);
        }

        public static string BuildValidReplacementString(string delimiter) => string.Format(ReplacementStringFormat, ReplacementPart, Regex.Escape(delimiter));
    }
}
