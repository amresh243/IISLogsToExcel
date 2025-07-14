namespace IISLogsToExcel
{
    public static class Extensons
    {
        public static int GetValidNumber(this string text)
        {
            if (int.TryParse(text, out int number))
                return number;

            return 0;
        }

        public static bool IsNumeric(this string input)
        {
            if (string.IsNullOrEmpty(input))
                return false;

            var nonDigit = input.Where(c => !char.IsDigit(c)).ToList();
            if (nonDigit.Count > 0)
                return false;

            return true;
        }
    }
}
