using SolidWorks.Interop.swconst;

namespace AngelSix.SolidDna
{
    /// <summary>
    /// Background processing options.
    /// <see cref="swBackgroundProcessOption_e"/>
    /// </summary>
    public enum BackgroundProcessOptions
    {
        // disabled
        BackgroundProcessingDisabled = swBackgroundProcessOption_e.swBackgroundProcessing_Disabled,
        // enabled
        BackgroundProcessingEnabled = swBackgroundProcessOption_e.swBackgroundProcessing_Enabled,
        // defer to ISldWorks::EnableBackgroundProcessing setting
        BackgroundProcessingDeferToApplication = swBackgroundProcessOption_e.swBackgroundProcessing_DeferToApplication
    }

}
