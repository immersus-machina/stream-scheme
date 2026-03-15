namespace StreamScheme.Examples.Models;

public record SalesReport
{
    public required string MostSoldItem { get; init; }
    public required string SecondMostSoldItem { get; init; }
    public required string HighestGrowthItem { get; init; }
    public required string LowestRevenueItem { get; init; }
    public required string MostReturnedItem { get; init; }
    public required Category MostSoldCategory { get; init; }
    public required Category FastestGrowingCategory { get; init; }
    public required Category LowestStockCategory { get; init; }
    public required Category MostDiscountedCategory { get; init; }
    public required Category HighestMarginCategory { get; init; }
}

public enum Category
{
    Electronics,
    Clothing,
    Food,
    Books,
    Sports,
    Home,
    Garden,
    Toys,
    Health,
    Beauty
}
