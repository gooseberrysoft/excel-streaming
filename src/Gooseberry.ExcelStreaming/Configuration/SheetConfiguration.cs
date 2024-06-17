// ReSharper disable once CheckNamespace

namespace Gooseberry.ExcelStreaming;

public readonly record struct SheetConfiguration(
    IReadOnlyCollection<Column>? Columns = null,
    uint? FreezeRows = null,
    uint? FreezeColumns = null,
    bool ShowGridLines = true);