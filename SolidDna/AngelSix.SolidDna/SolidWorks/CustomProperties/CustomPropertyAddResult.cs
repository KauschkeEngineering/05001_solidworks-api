using SolidWorks.Interop.swconst;

namespace AngelSix.SolidDna
{
    /// <summary>
    /// Result codes when adding custom properties.
    /// <see cref="swCustomInfoAddResult_e"/>
    /// </summary>
    public enum CustomPropertyAddResult
    {
        // success
        AddedOrChanged = swCustomInfoAddResult_e.swCustomInfoAddResult_AddedOrChanged,
        // failed to add the custom property
        GenericFail = swCustomInfoAddResult_e.swCustomInfoAddResult_GenericFail,
        // existing custom property with the same name has a different type
        MismatchAgainstExistingType = swCustomInfoAddResult_e.swCustomInfoAddResult_MismatchAgainstExistingType,
        // specified value of the custom property does not match the specified type
        MismatchAgainstSpecifiedType = swCustomInfoAddResult_e.swCustomInfoAddResult_MismatchAgainstSpecifiedType
    }
}
