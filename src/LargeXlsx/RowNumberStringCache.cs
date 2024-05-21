namespace LargeXlsx
{
    /// <summary>
    /// Caches a single row number string. Used to provide the CurrentRowNumber
    /// in string form during worksheet production where rows are processed
    /// consecutively and multiple columns generate string representations of the
    /// current row. Reduces small object heap allocations by a couple of GB in the
    /// Examples suite.
    /// </summary>
    internal sealed class RowNumberStringCache
    {
        private int _rowNumber;
        private string _rowNumberAsString;

        static RowNumberStringCache()
        {
        }

        private RowNumberStringCache()
        {
        }

        public static RowNumberStringCache Instance { get; } = new RowNumberStringCache();

        public string GetRowNumberAsString(int rowNumber)
        {
            if (_rowNumber == rowNumber)
            {
                return _rowNumberAsString;
            }

            _rowNumber = rowNumber;
            _rowNumberAsString = rowNumber.ToString();

            return _rowNumberAsString;
        }
    }
}
