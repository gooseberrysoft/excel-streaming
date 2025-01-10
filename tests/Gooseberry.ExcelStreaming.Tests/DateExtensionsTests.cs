using System.Globalization;
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

#if NET8_0_OR_GREATER
    [Fact]
    public void CheckCustomOADate()
    {
        var date = DateTime.Now.AddYears(-1500);

        while (date.Year != 2025)
        {
            double internalOADate = double.Parse(date.ToInternalOADate(), NumberStyles.Any, CultureInfo.InvariantCulture);

            double dateOADate = date.ToOADate();

            DateTime.FromOADate(internalOADate).Should().BeCloseTo(DateTime.FromOADate(dateOADate), TimeSpan.FromMilliseconds(1));

            date = date.AddDays(1).AddMicroseconds(Random.Shared.Next(1, 1_000));
        }
    }
#endif
}