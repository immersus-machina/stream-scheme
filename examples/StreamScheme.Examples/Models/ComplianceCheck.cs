namespace StreamScheme.Examples.Models;

public record ComplianceCheck
{
    public required int AuditId { get; init; }
    public ComplianceStatus? FireSafety { get; init; }
    public ComplianceStatus? ElectricalInspection { get; init; }
    public ComplianceStatus? PlumbingReview { get; init; }
    public ComplianceStatus? StructuralIntegrity { get; init; }
    public ComplianceStatus? ElevatorCertification { get; init; }
    public ComplianceStatus? HvacCompliance { get; init; }
    public ComplianceStatus? AccessibilityAudit { get; init; }
    public ComplianceStatus? EnvironmentalImpact { get; init; }
    public ComplianceStatus? NoiseCompliance { get; init; }
    public ComplianceStatus? WasteDisposal { get; init; }
    public ComplianceStatus? WaterQuality { get; init; }
    public ComplianceStatus? AirQuality { get; init; }
    public ComplianceStatus? PestControl { get; init; }
    public ComplianceStatus? EmergencyExits { get; init; }
    public ComplianceStatus? SignageCompliance { get; init; }
    public ComplianceStatus? ParkingRegulations { get; init; }
    public ComplianceStatus? ZoningCompliance { get; init; }
    public ComplianceStatus? InsuranceVerification { get; init; }
    public ComplianceStatus? OccupancyPermit { get; init; }
}

public enum ComplianceStatus
{
    CompliantNoIssuesFound,
    NonCompliantRemediationRequired,
    PendingInspectorReview,
    WaivedByBoardResolution,
    ExpiredRequiresRenewal
}
