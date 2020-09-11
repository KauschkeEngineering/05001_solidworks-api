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

    /// <summary>
    /// Result codes when setting custom properties.
    /// <see cref="swCustomInfoSetResult_e"/>
    /// </summary>
    public enum CustomPropertySetResult
    {
        // success
        OK = swCustomInfoSetResult_e.swCustomInfoSetResult_OK,
        // custom property does not exist
        NotPresent = swCustomInfoSetResult_e.swCustomInfoSetResult_NotPresent,
        // specified value has an incorrect typeSpecified value has an incorrect type
        TypeMismatch = swCustomInfoSetResult_e.swCustomInfoSetResult_TypeMismatch,
        // ?? not described in SOLIDWORKS API see: https://help.solidworks.com/2020/English/api/swconst/SOLIDWORKS.Interop.swconst~SOLIDWORKS.Interop.swconst.swCustomInfoSetResult_e.html?verRedirect=1
        // but this is returned if the property is linked for e.g. to parent.
        LinkedProp = swCustomInfoSetResult_e.swCustomInfoSetResult_LinkedProp
    }

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
        MismatchAgainstSpecifiedType = swCustomInfoAddResult_e.swCustomInfoAddResult_MismatchAgainstSpecifiedType,
    }

    /// <summary>
    /// Result codes when deleting custom properties.
    /// <see cref="swCustomInfoDeleteResult_e"/>
    /// </summary>
    public enum CustomPropertyDeleteResult
    {
        // success
        OK = swCustomInfoDeleteResult_e.swCustomInfoDeleteResult_OK,
        // custom property does not exist
        NotPresent = swCustomInfoDeleteResult_e.swCustomInfoDeleteResult_NotPresent,
        // ?? not described in SOLIDWORKS API see: http://help.solidworks.com/2020/english/api/swconst/SOLIDWORKS.Interop.swconst~SOLIDWORKS.Interop.swconst.swCustomInfoDeleteResult_e.html?verRedirect=1
        // but this is returned if the property is linked for e.g. to parent.
        LinkedProp = swCustomInfoDeleteResult_e.swCustomInfoDeleteResult_LinkedProp
    }

}
