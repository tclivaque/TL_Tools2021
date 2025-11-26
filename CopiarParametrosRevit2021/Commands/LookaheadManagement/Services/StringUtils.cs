using System.Linq;

namespace TL_Tools2021.Commands.LookaheadManagement.Services
{
    public static class StringUtils
    {
        public static string Clean(string input)
        {
            if (string.IsNullOrEmpty(input)) return "";
            char[] arr = input.Where(c => char.IsLetterOrDigit(c)).ToArray();
            return new string(arr).ToUpper();
        }
    }
}
