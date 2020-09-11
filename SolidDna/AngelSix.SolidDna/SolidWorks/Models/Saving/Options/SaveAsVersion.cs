using System;
using SolidWorks.Interop.swconst;

namespace AngelSix.SolidDna
{
    /// <summary>
    /// Version of a particular format to save the document, from <see cref="swSaveAsVersion_e"/>
    /// </summary>
    public enum SaveAsVersion
    {
        /// <summary>
        /// Saves the model in the default way with no special settings (this is the typical way)
        /// </summary>
        CurrentVersion = swSaveAsVersion_e.swSaveAsCurrentVersion,

        /// <summary>
        /// Obsolete and no longer supported
        /// </summary>
        [Obsolete("Use CurrentVersion instead")]
        SolidWorks98plus = swSaveAsVersion_e.swSaveAsSW98plus,

        /// <summary>
        /// Saves the model in Pro/E format
        /// </summary>
        FormatProE = swSaveAsVersion_e.swSaveAsFormatProE,

        /// <summary>
        /// Saves a detached drawing as a standard drawing
        /// </summary>
        StandardDrawing = swSaveAsVersion_e.swSaveAsStandardDrawing,

        /// <summary>
        /// Saves a standard drawing as a detached drawing
        /// </summary>
        DetachedDrawing = swSaveAsVersion_e.swSaveAsDetachedDrawing
    }
}
