namespace StreamScheme.OpenXml;

internal interface IOaDateConverter
{
    bool IsValidOaDate(double value);
    DateTime ToDateTime(double oaDate);
    double ToSerialDate(DateTime dateTime);
}

internal class OaDateConverter : IOaDateConverter
{
    private const double OaDateMin = -657435.0;
    private const double OaDateMax = 2958466.0;
    private static DateTime ExcelEpoch { get; } = new(1899, 12, 31, 0, 0, 0, DateTimeKind.Unspecified);

    public bool IsValidOaDate(double value) => value is > OaDateMin and < OaDateMax;


    public DateTime ToDateTime(double oaDate)
    {
        // Lotus 1-2-3 leap year bug: serial 60 = fictitious Feb 29, 1900.
        // When reading, serials >= 61 need -1 adjustment.
        var serial = (int)oaDate;
        var days = serial >= 61 ? serial - 1 : serial;
        var fractionalDay = oaDate - serial;
        return ExcelEpoch.AddDays(days).AddDays(fractionalDay);
    }

    public double ToSerialDate(DateTime dateTime)
    {
        // When writing, days >= 60 need +1 adjustment (inverse of reading).
        var days = (int)(dateTime.Date - ExcelEpoch).TotalDays;
        var serial = days >= 60 ? days + 1 : days;
        var fractionalDay = dateTime.TimeOfDay.TotalDays;
        return serial + fractionalDay;
    }
}
