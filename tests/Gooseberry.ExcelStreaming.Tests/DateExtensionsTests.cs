using FluentAssertions;
using Gooseberry.ExcelStreaming.Extensions;
using Gooseberry.ExcelStreaming.Writers;
using Xunit;

namespace Gooseberry.ExcelStreaming.Tests;

public sealed class DateExtensionsTests
{
    [Fact]
    public void CheckDateOnlyOADate()
    {
        var date = new DateOnly(999, 01, 01);

        while (date.Year != 2025)
        {
            double dateOnlyOADate = Convert.ToDouble(date.ToOADate());

            double dateOADate = date.ToDateTime(TimeOnly.MinValue).ToOADate();

            dateOnlyOADate.Should().Be(dateOADate);

            date = date.AddDays(1);
        }
    }
}