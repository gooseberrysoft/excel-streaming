using System.Collections.Generic;
using System.Runtime.InteropServices;

namespace Gooseberry.ExcelStreaming.Configuration
{
    [StructLayout(LayoutKind.Auto)]
    public readonly struct SheetConfiguration
    {
        public SheetConfiguration(
            IReadOnlyCollection<Column>? columns = null, 
            IReadOnlyCollection<Merge>? merges = null, 
            CellReference? topLeftUnpinnedCell = null)
        {
            Columns = columns;
            Merges = merges;
            TopLeftUnpinnedCell = topLeftUnpinnedCell;
        }

        public IReadOnlyCollection<Column>? Columns { get; }
        
        public IReadOnlyCollection<Merge>? Merges { get; }
        
        public CellReference? TopLeftUnpinnedCell { get; }
    }
}