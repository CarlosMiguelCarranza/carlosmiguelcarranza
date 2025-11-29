
namespace cmc.ExcelUtilities
{
    [Serializable]
    internal class ExcelReaderException : Exception
    {
        public ExcelReaderException()
        {
        }

        public ExcelReaderException(string? message) : base(message)
        {
        }

        public ExcelReaderException(string? message, Exception? innerException) : base(message, innerException)
        {
        }
    }
}