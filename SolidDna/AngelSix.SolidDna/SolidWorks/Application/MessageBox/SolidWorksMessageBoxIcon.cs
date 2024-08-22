using SolidWorks.Interop.swconst;

namespace AngelSix.SolidDna
{
    /// <summary>
    /// The type of message box icon for a SolidWorks message, of type <see cref="swMessageBoxIcon_e"/>
    /// </summary>
    public enum SolidWorksMessageBoxIcon
    {
        /// <summary>
        /// A warning icon
        /// </summary>
        Warning = swMessageBoxIcon_e.swMbWarning,

        /// <summary>
        /// An information icon
        /// </summary>
        Information = swMessageBoxIcon_e.swMbInformation,

        /// <summary>
        /// A question mark icon
        /// </summary>
        Question = swMessageBoxIcon_e.swMbQuestion,

        /// <summary>
        /// An exclamation icon
        /// </summary>
        Stop = swMessageBoxIcon_e.swMbStop
    }
}
