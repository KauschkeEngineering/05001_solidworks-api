using SolidWorks.Interop.swconst;

namespace AngelSix.SolidDna
{
    /// <summary>
    /// Suppression errors.
    /// <see cref="swSuppressionError_e"/>
    /// </summary>
    public enum ComponentSuppressionErrors
    {
        // component object is no longer valid; for example, if a configuration changed
        BadComponent = swSuppressionError_e.swSuppressionBadComponent,
        // invalid state was specified
        BadState = swSuppressionError_e.swSuppressionBadState,
        // state was changed
        ChangeOk = swSuppressionError_e.swSuppressionChangeOk,
        // change failed, even though the arguments were okay
        ChangeFailed = swSuppressionError_e.swSuppressionChangeFailed
    }

}
