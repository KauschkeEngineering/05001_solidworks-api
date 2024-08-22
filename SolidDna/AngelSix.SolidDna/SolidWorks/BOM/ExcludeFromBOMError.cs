using SolidWorks.Interop.swconst;

namespace AngelSix.SolidDna
{
    /// <summary>
    /// Result codes when setting whether to exclude a component from a bill of materials.
    /// <see cref="swExcludeFromBOMError_e"/>
    /// </summary>
    public enum ExcludeFromBOMError
    {
        Fail = swExcludeFromBOMError_e.swExcludeFromBOM_Fail,
        Success = swExcludeFromBOMError_e.swExcludeFromBOM_Success,
        EnvelopedComponent = swExcludeFromBOMError_e.swExcludeFromBOM_EnvelopedComponent
    }
}
