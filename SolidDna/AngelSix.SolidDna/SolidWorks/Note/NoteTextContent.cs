using SolidWorks.Interop.swconst;

namespace AngelSix.SolidDna
{
    /// <summary>
    /// BOM balloon note text styles.
    /// <see cref="swDetailingNoteTextContent_e"/>
    /// </summary>
    public enum NoteTextContent
    {
        /// <summary>
        /// Custom text
        /// </summary>
        Custom = swDetailingNoteTextContent_e.swDetailingNoteTextCustom,

        /// <summary>
        /// Custom property
        /// </summary>
        CustomProperty = swDetailingNoteTextContent_e.swDetailingNoteTextCustomProperty,

        /// <summary>
        /// Item number
        /// </summary>
        ItemNumber = swDetailingNoteTextContent_e.swDetailingNoteTextItemNumber,

        /// <summary>
        /// Text quantity
        /// </summary>
        TextQuantity = swDetailingNoteTextContent_e.swDetailingNoteTextQuantity
    }
}
