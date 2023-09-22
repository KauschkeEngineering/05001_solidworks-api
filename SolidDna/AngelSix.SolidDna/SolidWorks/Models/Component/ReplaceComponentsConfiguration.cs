using SolidWorks.Interop.swconst;

namespace AngelSix.SolidDna
{
    /// <summary>
    /// Suppression errors.
    /// <see cref="swSuppressionError_e"/>
    /// </summary>
    public enum ReplaceComponentsConfiguration
    {
        // Let SOLIDWORKS attempt to match the configuration of the old components with a configuration in the replacement component
        MatchName = swReplaceComponentsConfiguration_e.swReplaceComponentsConfiguration_MatchName,
        // Use the specified configuration
        ManuallySelect = swReplaceComponentsConfiguration_e.swReplaceComponentsConfiguration_ManuallySelect
    }

}
