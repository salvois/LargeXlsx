namespace LargeXlsx
{
    public class XlsxDataIgnoreError
    {
        internal XlsxDataIgnoreError(string address, ErrorType ignoreErrorType)
        {
            IgnoreErrorType = ignoreErrorType;
            this.Address = address;
        }

        public string Address { get; }

        public ErrorType IgnoreErrorType { get; }
        
        public enum ErrorType
        {
            NumberStoredAsText
        }
    }
}