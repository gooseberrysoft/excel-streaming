using System.Globalization;
using FluentAssertions;
using Gooseberry.ExcelStreaming.Extensions;
using Xunit;

namespace Gooseberry.ExcelStreaming.Tests;

public sealed class DateExtensionsTests
{
    [Fact]
    public void CheckDateOnlyOADate()
    {
        var date = new DateOnly(100, 01, 01);

        while (date.Year != 2100)
        {
            double dateOnlyOADate = Convert.ToDouble(date.ToOADate());

            double dateOADate = date.ToDateTime(TimeOnly.MinValue).ToOADate();

            dateOnlyOADate.Should().Be(dateOADate);

            date = date.AddDays(1);
        }
    }

#if NET8_0_OR_GREATER
    [Fact]
    public void CheckCustomOADateByDay()
    {
        var date = DateTime.Now.AddYears(-1500);

        while (date.Year != 2025)
        {
            double internalOADate = double.Parse(date.ToInternalOADate(), NumberStyles.Any, CultureInfo.InvariantCulture);

            double dateOADate = date.ToOADate();

            DateTime.FromOADate(internalOADate).Should().BeCloseTo(DateTime.FromOADate(dateOADate), TimeSpan.FromMilliseconds(1));

            date = date.AddDays(1).AddMinutes(25).AddMicroseconds(Random.Shared.Next(1, 1_000));
        }
    }

    [Fact]
    public void CheckCustomOADateByMinutes()
    {
        var date = DateTime.Now;

        while (date.Day != DateTime.Now.AddDays(20).Day)
        {
            double internalOADate = double.Parse(date.ToInternalOADate(), NumberStyles.Any, CultureInfo.InvariantCulture);

            double dateOADate = date.ToOADate();

            DateTime.FromOADate(internalOADate).Should().BeCloseTo(DateTime.FromOADate(dateOADate), TimeSpan.FromMilliseconds(1));

            date = date.AddMinutes(1);
        }
    }

    [Fact]
    public void CheckCustomOADateByMilliseconds()
    {
        var date = DateTime.Now;

        while (date.Day != DateTime.Now.AddDays(2).Day)
        {
            double internalOADate = double.Parse(date.ToInternalOADate(), NumberStyles.Any, CultureInfo.InvariantCulture);

            double dateOADate = date.ToOADate();

            DateTime.FromOADate(internalOADate).Should().BeCloseTo(DateTime.FromOADate(dateOADate), TimeSpan.FromMilliseconds(1));

            date = date.AddMilliseconds(100);
        }
    }
#endif
}