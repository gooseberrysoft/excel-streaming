// ReSharper disable once CheckNamespace

namespace Gooseberry.ExcelStreaming;

public record SheetConfiguration(
    IReadOnlyCollection<Column>? Columns = null,
    uint? FrozenRows = null,
    uint? FrozenColumns = null,
    bool ShowGridLines = true,
    CellRange? AutoFilter = null);