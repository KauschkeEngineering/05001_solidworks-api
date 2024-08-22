using SolidWorks.Interop.swconst;

namespace AngelSix.SolidDna
{
    /// <summary>
    /// The type of message box buttons for a SolidWorks message, of type <see cref="swMessageBoxBtn_e"/>
    /// </summary>
    public enum SolidWorksMessageBoxButtons
    {
        /// <summary>
        /// An Abort, Retry and Ignore button
        /// </summary>
        AbortRetryIgnore = swMessageBoxBtn_e.swMbAbortRetryIgnore,

        /// <summary>
        /// A single OK button
        /// </summary>
        Ok = swMessageBoxBtn_e.swMbOk,

        /// <summary>
        /// An OK and Cancel button
        /// </summary>
        OkCancel = swMessageBoxBtn_e.swMbOkCancel,

        /// <summary>
        /// A single Retry button
        /// </summary>
        RetryCancel = swMessageBoxBtn_e.swMbRetryCancel,

        /// <summary>
        /// A Yes and No button
        /// </summary>
        YesNo = swMessageBoxBtn_e.swMbYesNo,

        /// <summary>
        /// A Yes, No and Cancel button
        /// </summary>
        YesNoCancel = swMessageBoxBtn_e.swMbYesNoCancel
    }
}
