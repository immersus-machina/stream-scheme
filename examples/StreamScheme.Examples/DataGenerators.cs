using StreamScheme;
using StreamScheme.Examples.Models;

namespace StreamScheme.Examples;

public static class DataGenerators
{
    private static Category[] AllCategories { get; } = Enum.GetValues<Category>();
    private static ComplianceStatus[] AllStatuses { get; } = Enum.GetValues<ComplianceStatus>();

    public static SalesReport[] GenerateSalesReports(int count)
    {
        var reports = new SalesReport[count];

        for (var row = 0; row < count; row++)
        {
            reports[row] = new SalesReport
            {
                MostSoldItem = $"SKU-{row + 1}-best".PadRight(20, 'x'),
                SecondMostSoldItem = $"SKU-{row + 1}-second".PadRight(20, 'x'),
                HighestGrowthItem = $"SKU-{row + 1}-growth".PadRight(20, 'x'),
                LowestRevenueItem = $"SKU-{row + 1}-low".PadRight(20, 'x'),
                MostReturnedItem = $"SKU-{row + 1}-return".PadRight(20, 'x'),
                MostSoldCategory = AllCategories[(row + 0) % AllCategories.Length],
                FastestGrowingCategory = AllCategories[(row + 1) % AllCategories.Length],
                LowestStockCategory = AllCategories[(row + 2) % AllCategories.Length],
                MostDiscountedCategory = AllCategories[(row + 3) % AllCategories.Length],
                HighestMarginCategory = AllCategories[(row + 4) % AllCategories.Length],
            };
        }

        return reports;
    }

    public static ComplianceCheck[] GenerateComplianceChecks(int count, double fillRate = 0.30)
    {
        var random = new Random(42);
        var checks = new ComplianceCheck[count];

        for (var row = 0; row < count; row++)
        {
            checks[row] = new ComplianceCheck
            {
                AuditId = row + 1,
                FireSafety = MaybeStatus(random, fillRate),
                ElectricalInspection = MaybeStatus(random, fillRate),
                PlumbingReview = MaybeStatus(random, fillRate),
                StructuralIntegrity = MaybeStatus(random, fillRate),
                ElevatorCertification = MaybeStatus(random, fillRate),
                HvacCompliance = MaybeStatus(random, fillRate),
                AccessibilityAudit = MaybeStatus(random, fillRate),
                EnvironmentalImpact = MaybeStatus(random, fillRate),
                NoiseCompliance = MaybeStatus(random, fillRate),
                WasteDisposal = MaybeStatus(random, fillRate),
                WaterQuality = MaybeStatus(random, fillRate),
                AirQuality = MaybeStatus(random, fillRate),
                PestControl = MaybeStatus(random, fillRate),
                EmergencyExits = MaybeStatus(random, fillRate),
                SignageCompliance = MaybeStatus(random, fillRate),
                ParkingRegulations = MaybeStatus(random, fillRate),
                ZoningCompliance = MaybeStatus(random, fillRate),
                InsuranceVerification = MaybeStatus(random, fillRate),
                OccupancyPermit = MaybeStatus(random, fillRate),
            };
        }

        return checks;
    }

    public static string[][] GenerateUniqueStrings(int rowCount, int columnCount = 10)
    {
        string[] randomSuffixes =
        [
            " <div>",
            " just me",
            " \"quoted\"",
            " naive cafe",
            " test",
            " text",
            " price 100",
            " a&b<c>d",
            " resume",
            " foo\tbar"
        ];

        var rows = new string[rowCount][];

        for (var row = 0; row < rowCount; row++)
        {
            var cells = new string[columnCount];
            for (var col = 0; col < columnCount; col++)
            {
                cells[col] = $"R{row + 1}C{col + 1}{randomSuffixes[col % randomSuffixes.Length]}";
            }

            rows[row] = cells;
        }

        return rows;
    }

    public static IEnumerable<FieldValue[]> GenerateMixedRows(int count)
    {
        var baseDate = new DateOnly(2020, 1, 1);
        for (var i = 0; i < count; i++)
        {
            yield return
            [
                (FieldValue)$"Name-{i}",
                (FieldValue)$"Category-{i % 10}",
                (FieldValue)(double)i,
                (FieldValue)(i * 1.5),
                (FieldValue)baseDate.AddDays(i % 365),
                (FieldValue)(i % 2 == 0),
            ];
        }
    }

    private static ComplianceStatus? MaybeStatus(Random random, double fillRate) =>
        random.NextDouble() < fillRate
            ? AllStatuses[random.Next(AllStatuses.Length)]
            : null;
}
