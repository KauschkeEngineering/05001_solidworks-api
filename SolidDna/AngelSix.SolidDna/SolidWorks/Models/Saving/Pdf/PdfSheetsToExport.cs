using SolidWorks.Interop.swconst;

namespace AngelSix.SolidDna
{
    /// <summary>
    /// Export data sheets to export options.
    /// Specifies which drawing sheets will get exported to PDF, from <see cref="SolidWorks.Interop.swconst.swExportDataSheetsToExport_e"/>
    /// </summary>
    public enum PdfSheetsToExport
    {
        /// <summary>
        /// Exports all drawing sheets
        /// </summary>
        ExportAllSheets = swExportDataSheetsToExport_e.swExportData_ExportAllSheets,

        /// <summary>
        /// Exports the currently active sheet
        /// </summary>
        ExportCurrentSheet = swExportDataSheetsToExport_e.swExportData_ExportCurrentSheet,

        /// <summary>
        /// Exports the sheets specified in the sheets array
        /// </summary>
        ExportSpecifiedSheets = swExportDataSheetsToExport_e.swExportData_ExportSpecifiedSheets
    }
}
