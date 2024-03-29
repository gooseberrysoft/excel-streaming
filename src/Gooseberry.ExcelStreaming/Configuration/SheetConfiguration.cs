using System.Collections.Generic;
using System.Runtime.InteropServices;

namespace Gooseberry.ExcelStreaming.Configuration
{
    [StructLayout(LayoutKind.Auto)]
    public readonly struct SheetConfiguration
    {
        public SheetConfiguration(
            IReadOnlyCollection<Column>? columns = null, 
            CellReference? topLeftUnpinnedCell = null,
            bool showGridLines = true)
        {
            ShowGridLines = showGridLines;
            Columns = columns;
            TopLeftUnpinnedCell = topLeftUnpinnedCell;
        }

        public IReadOnlyCollection<Column>? Columns { get; }
        
        public CellReference? TopLeftUnpinnedCell { get; }
        
        public bool ShowGridLines { get; }
    }
}