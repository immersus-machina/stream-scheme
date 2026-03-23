namespace StreamScheme.FSharp.DIExample

open StreamScheme.FSharp

/// File path for the sales report spreadsheet.
type SalesReportPath = SalesReportPath of string

/// Product sales data to export.
type ProductSalesData =
    { Header: FieldValue seq
      Rows: FieldValue seq seq }
