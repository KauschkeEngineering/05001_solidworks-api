using SolidWorks.Interop.swconst;

namespace AngelSix.SolidDna
{
    /// <summary>
    /// The result for a SolidWorks message, of type <see cref="swMessageBoxResult_e"/>
    /// </summary>
    public enum SolidWorksMessageBoxResult
    {
        /// <summary>
        /// Abort was clicked
        /// </summary>
        Abort = swMessageBoxResult_e.swMbHitAbort,

        /// <summary>
        /// Ignore was clicked
        /// </summary>
        Ignore = swMessageBoxResult_e.swMbHitIgnore,

        /// <summary>
        /// No was clicked
        /// </summary>
        No = swMessageBoxResult_e.swMbHitNo,

        /// <summary>
        /// Ok was clicked
        /// </summary>
        Ok = swMessageBoxResult_e.swMbHitOk,

        /// <summary>
        /// Retry was clicked
        /// </summary>
        Retry = swMessageBoxResult_e.swMbHitRetry,

        /// <summary>
        /// Yes was clicked
        /// </summary>
        Yes = swMessageBoxResult_e.swMbHitYes,

        /// <summary>
        /// Cancel was clicked
        /// </summary>
        Cancel = swMessageBoxResult_e.swMbHitCancel
    }
}
