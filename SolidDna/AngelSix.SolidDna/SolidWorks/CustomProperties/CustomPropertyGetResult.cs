using SolidWorks.Interop.swconst;

namespace AngelSix.SolidDna
{
    /// <summary>
    /// Result codes when getting custom properties.
    /// <see cref="swCustomInfoGetResult_e"/>
    /// </summary>
    public enum CustomPropertyGetResult
    {
        // cached value was returned
        CachedValue = swCustomInfoGetResult_e.swCustomInfoGetResult_CachedValue,
        // custom property does not exist
        NotPresent = swCustomInfoGetResult_e.swCustomInfoGetResult_NotPresent,
        // resolved value was returned
        ResolvedValue = swCustomInfoGetResult_e.swCustomInfoGetResult_ResolvedValue
    }
}
