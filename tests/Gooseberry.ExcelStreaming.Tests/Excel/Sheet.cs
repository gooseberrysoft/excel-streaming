using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using Gooseberry.ExcelStreaming.Configuration;

namespace Gooseberry.ExcelStreaming.Tests.Excel
{
    [StructLayout(LayoutKind.Auto)]
    public readonly struct Sheet : IEquatable<Sheet>
    {
        public Sheet(string name, IReadOnlyCollection<Row> rows, IReadOnlyCollection<Column>? columns = null)
        {
            Name = name;
            Columns = columns ?? Array.Empty<Column>();
            Rows = rows;
        }

        public string Name { get; }

        public IReadOnlyCollection<Column> Columns { get; }

        public IReadOnlyCollection<Row> Rows { get; }

        public bool Equals(Sheet other)
            => string.Equals(Name, other.Name, StringComparison.Ordinal) && Rows.SequenceEqual(other.Rows);

        public override bool Equals(object? other)
            => other is Sheet sheet && Equals(sheet);

        public override int GetHashCode()
            => HashCode.Combine(Name, Rows);
    }
}
