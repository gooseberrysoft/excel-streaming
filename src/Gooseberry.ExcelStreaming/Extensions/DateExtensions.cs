using System.Runtime.CompilerServices;

namespace Gooseberry.ExcelStreaming.Extensions;

internal static class DateExtensions
{
    //Constants from DateTime implementation
    // Number of days in a non-leap year
    private const int DaysPerYear = 365;

    // Number of days in 4 years
    private const int DaysPer4Years = DaysPerYear * 4 + 1; // 1461

    // Number of days in 100 years
    private const int DaysPer100Years = DaysPer4Years * 25 - 1; // 36524

    // Number of days in 400 years
    private const int DaysPer400Years = DaysPer100Years * 4 + 1; // 146097

    // Number of days from 1/1/0001 to 12/30/1899
    private const int DaysTo1899 = DaysPer400Years * 4 + DaysPer100Years * 3 - 367;

    //we don't use DateTime.ToOADate because writing double to string extremely slow 
    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public static int ToOADate(this DateOnly value)
        => value.DayNumber - DaysTo1899;
}