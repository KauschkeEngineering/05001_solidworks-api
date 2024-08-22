using SolidWorks.Interop.swconst;

namespace AngelSix.SolidDna
{
    /// <summary>
    /// Style of the text contents of balloon notes
    /// </summary>
    public enum BalloonTextContent
    {
        /// <summary>
        /// Use the document layout default
        /// </summary>
        DocumentDefault = -1,

        /// <summary>
        /// Custom text
        /// </summary>
        Custom = swBalloonTextContent_e.swBalloonTextCustom,

        /// <summary>
        /// Item number
        /// </summary>
        ItemNumber = swBalloonTextContent_e.swBalloonTextItemNumber,

        /// <summary>
        /// Quantity
        /// </summary>
        Quantity = swBalloonTextContent_e.swBalloonTextQuantity,

        /// <summary>
        /// Custom properties
        /// </summary>
        CustomProperties = swBalloonTextContent_e.swBalloonTextCustomProperties,

        /// <summary>
        /// Component reference
        /// </summary>
        ComponentReference = swBalloonTextContent_e.swBalloonTextComponentReference,

        /// <summary>
        /// Spool reference
        /// </summary>
        SpoolReference = swBalloonTextContent_e.swBalloonTextSpoolReference,

        /// <summary>
        /// Part number BOM
        /// </summary>
        PartNumberBOM = swBalloonTextContent_e.swBalloonTextPartNumberBOM,

        /// <summary>
        /// File name
        /// </summary>
        FileName = swBalloonTextContent_e.swBalloonTextFileName,

        /// <summary>
        /// Cutlist properties
        /// </summary>
        CutlistProperties = swBalloonTextContent_e.swBalloonTextCutlistProperties,

        /// <summary>
        /// View sheet
        /// </summary>
        ViewSheet = swBalloonTextContent_e.swBalloonTextViewSheet,

        /// <summary>
        /// View sheet with label
        /// </summary>
        ViewSheetWithLabel = swBalloonTextContent_e.swBalloonTextViewSheetWithLabel,

        /// <summary>
        /// View zone
        /// </summary>
        ViewZone = swBalloonTextContent_e.swBalloonTextViewZone,

        /// <summary>
        /// View letter
        /// </summary>
        ViewLetter = swBalloonTextContent_e.swBalloonTextViewViewLetter
    }
}
