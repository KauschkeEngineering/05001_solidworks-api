using SolidWorks.Interop.swconst;

namespace AngelSix.SolidDna
{
    /// <summary>
    /// The opens for a document type used in calls such as <see cref="SolidWorksApplication.OpenFile(string, bool)"/>
    /// from <see cref="swDocumentTypes_e"/>
    /// </summary>
    public enum DocumentType
    {
        /// <summary>
        /// Unknown
        /// </summary>
        None = swDocumentTypes_e.swDocNONE,
        /// <summary>
        /// SolidWorks Part
        /// </summary>
        Part = swDocumentTypes_e.swDocPART,
        /// <summary>
        /// SolidWorks Assembly
        /// </summary>
        Assembly = swDocumentTypes_e.swDocASSEMBLY,
        /// <summary>
        /// SolidWorks Drawing
        /// </summary>
        Drawing = swDocumentTypes_e.swDocDRAWING,
        /// <summary>
        /// SolidWorks Document Manager File
        /// </summary>
        DocumentManager = swDocumentTypes_e.swDocSDM,
        /// <summary>
        /// External File
        /// </summary>
        ExternalFile = swDocumentTypes_e.swDocLAYOUT,
        /// <summary>
        /// Imported Part
        /// </summary>
        ImportedPart = swDocumentTypes_e.swDocIMPORTED_PART,
        /// <summary>
        /// Imported Assembly
        /// </summary>
        ImportedAssembly = swDocumentTypes_e.swDocIMPORTED_ASSEMBLY,
    }
}
