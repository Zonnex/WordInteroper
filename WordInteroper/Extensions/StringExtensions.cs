using System;

namespace WordInteroper.Extensions
{
    public static class StringExtensions
    {
        public static bool Contains(this string source, string substring, StringComparison comparer)
        {
            return source?.IndexOf(substring, comparer) >= 0;
        }

        public static bool HasValue(this string source)
        {
            return !string.IsNullOrWhiteSpace(source);
        }
    }
}