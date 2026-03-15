namespace StreamScheme.OpenXml;

internal interface IOaDateConverter
{
    bool IsValidOaDate(double value);
    DateOnly ToDateOnly(double oaDate);
    int ToSerialDate(DateOnly date);
}

internal class OaDateConverter : IOaDateConverter
{
    private const double OaDateMin = -657435.0;
    private const double OaDateMax = 2958466.0;
    private static DateOnly ExcelEpoch { get; } = new(1899, 12, 31);

    public bool IsValidOaDate(double value) => value is > OaDateMin and < OaDateMax;

    // Lotus 1-2-3 leap year bug: serial 60 = fictitious Feb 29, 1900.
    // When reading, serials >= 61 need -1 adjustment.
    public DateOnly ToDateOnly(double oaDate)
    {
        var serial = (int)oaDate;
        var days = serial >= 61 ? serial - 1 : serial;
        return ExcelEpoch.AddDays(days);
    }

    // When writing, days >= 60 need +1 adjustment (inverse of reading).
    public int ToSerialDate(DateOnly date)
    {
        var days = date.DayNumber - ExcelEpoch.DayNumber;
        return days >= 60 ? days + 1 : days;
    }
}
