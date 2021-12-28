using System.Runtime.InteropServices;

namespace Gooseberry.ExcelStreaming.Configuration
{
    [StructLayout(LayoutKind.Auto)]
    public readonly struct CellReference
    {
        public CellReference(string column, int row)
        {
            // ToDo check column and row
            Value = $"{column}{row}";
        }
        
        public string Value { get; }
    }
}