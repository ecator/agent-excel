namespace AgentExcel.Models;

/// <summary>
/// Connection details of a shape if it is a connector.
/// </summary>
/// <param name="BeginShapeName">The name of the shape connected at the beginning of the connector.</param>
/// <param name="BeginConnectionSite">The connection site index on the beginning connected shape.</param>
/// <param name="EndShapeName">The name of the shape connected at the end of the connector.</param>
/// <param name="EndConnectionSite">The connection site index on the end connected shape.</param>
public record ShapeConnectionInfo(
    string? BeginShapeName,
    int? BeginConnectionSite,
    string? EndShapeName,
    int? EndConnectionSite
);
