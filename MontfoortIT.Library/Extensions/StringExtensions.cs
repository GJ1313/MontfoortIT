using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace MontfoortIT.Library.Extensions
{
    public static class StringExtention
    {
        public static bool ContainsIgnoreCase(this string source, string toCheck)
        {
            if (string.IsNullOrEmpty(toCheck) || string.IsNullOrEmpty(source))
                return true;

            return source.IndexOf(toCheck, StringComparison.InvariantCultureIgnoreCase) >= 0;
        }


        public static string CutOut(this string source, string valueToCutOut)
        {
            if (string.IsNullOrEmpty(source))
                return source;

            StringBuilder resultBuilder = new StringBuilder(source);            
            string sourceToLower = source.ToLower();
            var valueToCutOutLow = valueToCutOut.ToLower();
            int indexOfCut = sourceToLower.IndexOf(valueToCutOutLow);

            int indxOffset=0;
            while (indexOfCut >= 0)
            {
                resultBuilder = resultBuilder.Remove(indexOfCut + indxOffset, valueToCutOut.Length);
                indxOffset -= valueToCutOut.Length;
                                
                indexOfCut = sourceToLower.IndexOf(valueToCutOutLow, indexOfCut+1);
            }

            return resultBuilder.ToString();
        }

        public static string CutOut(this string source, params char[] valuesToCutOut)
        {
            if (string.IsNullOrEmpty(source))
                return source;

            StringBuilder resultBuilder = new StringBuilder(source);
            string sourceToLower = source.ToLower();
            var valueToCutOutLow = valuesToCutOut.Select(c=>char.ToLower(c)).ToArray();
            int indexOfCut = sourceToLower.IndexOfAny(valuesToCutOut);

            int indxOffset = 0;
            while (indexOfCut >= 0)
            {
                resultBuilder = resultBuilder.Remove(indexOfCut + indxOffset, 1);
                indxOffset --;

                indexOfCut = sourceToLower.IndexOfAny(valuesToCutOut, indexOfCut + 1);
            }

            return resultBuilder.ToString();
        }



        public static string ReplaceIgnoreCase(this string source, string valueToReplace, string valueToReplaceWith)
        {
            if (string.IsNullOrEmpty(source))
                return source;

            string sourceToLower = source.ToLower();

            int indexOfReplace = sourceToLower.IndexOf(valueToReplace.ToLower());
            string result = source;
            while (indexOfReplace >= 0)
            {
                result = result.Substring(0, indexOfReplace) + valueToReplaceWith + result.Substring(indexOfReplace + valueToReplace.Length);
                indexOfReplace = result.ToLower().IndexOf(valueToReplace.ToLower());
            }
            return result;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="source"></param>
        /// <param name="replaceValues">The last value is the value what is the replace value</param>
        /// <returns></returns>
        public static string MultiReplace(this string source, char valToReplaceWith, params char[] replaceValues)
        {
            StringBuilder builder = new StringBuilder();            
            foreach (var sourceChar in source)
            {
                if (replaceValues.Contains(sourceChar))
                    builder.Append(valToReplaceWith);
                else
                    builder.Append(sourceChar);
            }

            return builder.ToString();
        }
        
        public static string CleanSpaces(this string val)
        {
            StringBuilder stringBuilder = new StringBuilder();
            char? lastWhiteSpace = null;
            foreach (var c in val)
            {
                if (char.IsWhiteSpace(c))
                {
                    lastWhiteSpace = c;
                }
                else
                {
                    if (lastWhiteSpace.HasValue)
                    {
                        stringBuilder.Append(lastWhiteSpace);
                        lastWhiteSpace = null;
                    }

                    stringBuilder.Append(c);
                }
            }

            return stringBuilder.ToString().Trim();
        }


        /// <summary>
        /// Remove white spaces and replaces these by space
        /// </summary>
        /// <param name="val"></param>
        /// <returns></returns>
        public static string CleanWhiteSpace(this string val)
        {
            return val.Replace('\t', ' ').Replace('\r', ' ').Replace('\n', ' ').CleanSpaces();
        }



        /// <summary>
        /// Tries each seperator and when found a match use that one for the split / ignore the rest
        /// </summary>
        /// <param name="val"></param>
        /// <param name="sepeartors"></param>
        /// <returns></returns>
        public static string[] SplitConditional(this string val, params string[] sepeartors)
        {
            foreach (var seperator in sepeartors)
            {
                if (val.Contains(seperator))
                    return val.Split(new[] { seperator }, StringSplitOptions.RemoveEmptyEntries);
            }

            return new[] { val };
        }

        public static string Between(this string val, char seperator)
        {
            if (string.IsNullOrEmpty(val))
                return null;

            int startIndex = val.IndexOf(seperator);
            int endIndex = val.LastIndexOf(seperator);
            if (startIndex == -1 || endIndex == -1)
                return null;

            return val.Substring(startIndex + 1, endIndex - startIndex - 1);
        }

    }
}
