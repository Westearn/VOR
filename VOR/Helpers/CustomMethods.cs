using System;
using System.Collections.Generic;

namespace VOR.Helpers
{
    public static class CustomMethods
    {
        public static bool ContainsCustom(this string source, string target, StringComparison comparison = StringComparison.Ordinal)
        {
            if (string.IsNullOrEmpty(source) || string.IsNullOrEmpty(target))
            {
                return false;
            }

            string normalizedSource = NormalizeStrings(source);
            string normalizedTarget = NormalizeStrings(target);

            return normalizedSource.IndexOf(normalizedTarget, comparison) >= 0;
        }

        private static string NormalizeStrings(string input)
        {
            Dictionary<char, char> replacementMap = new Dictionary<char, char>()
            {
                { 'A', 'А' }, { 'E', 'Е' }, { 'O', 'О' },
                { 'P', 'Р' }, { 'C', 'С' }, { 'X', 'Х' },
                { 'a', 'а' }, { 'e', 'е' }, { 'o', 'о' },
                { 'p', 'р' }, { 'c', 'с' }, { 'x', 'х' }
            };

            char[] normalizedChars = input.ToCharArray();

            for (int i = 0; i < normalizedChars.Length; i++)
            {
                if (replacementMap.ContainsKey(normalizedChars[i]))
                {
                    normalizedChars[i] = replacementMap[normalizedChars[i]];
                }
            }

            return new string(normalizedChars);
        }

        public static string NormalizeString(this string source)
        {
            Dictionary<char, char> replacementMap = new Dictionary<char, char>()
            {
                { 'A', 'А' }, { 'E', 'Е' }, { 'O', 'О' },
                { 'P', 'Р' }, { 'C', 'С' }, { 'X', 'Х' },
                { 'a', 'а' }, { 'e', 'е' }, { 'o', 'о' },
                { 'p', 'р' }, { 'c', 'с' }, { 'x', 'х' }
            };

            char[] normalizedChars = source.ToCharArray();

            for (int i = 0; i < normalizedChars.Length; i++)
            {
                if (replacementMap.ContainsKey(normalizedChars[i]))
                {
                    normalizedChars[i] = replacementMap[normalizedChars[i]];
                }
            }

            return new string(normalizedChars);
        }
    }
}
