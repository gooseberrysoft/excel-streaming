using System.Collections.Generic;
using System.Runtime.InteropServices;

namespace Gooseberry.ExcelStreaming.Configuration
{
    [StructLayout(LayoutKind.Auto)]
    public readonly struct SheetConfiguration
    {
        public SheetConfiguration(
            IReadOnlyCollection<Column>? columns = null, 
            CellReference? topLeftUnpinnedCell = null)
        {
            Columns = columns;
            TopLeftUnpinnedCell = topLeftUnpinnedCell;
        }

        public IReadOnlyCollection<Column>? Columns { get; }
        
        public CellReference? TopLeftUnpinnedCell { get; }
    }
}