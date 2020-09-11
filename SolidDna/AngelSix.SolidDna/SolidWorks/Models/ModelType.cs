using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

namespace AngelSix.SolidDna
{
    /// <summary>
    /// Document types.
    /// The type of SolidWorks model the <see cref="ModelDoc2"/> is, from <see cref="swDocumentTypes_e"/>
    /// </summary>
    public enum ModelType
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
