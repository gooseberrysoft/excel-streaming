// ReSharper disable once CheckNamespace

namespace Gooseberry.ExcelStreaming;

public readonly record struct SheetConfiguration(
    IReadOnlyCollection<Column>? Columns = null,
    uint? FrozenRows = null,
    uint? FrozenColumns = null,
    bool ShowGridLines = true);