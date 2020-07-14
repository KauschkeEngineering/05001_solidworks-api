using SolidWorks.Interop.swconst;

namespace AngelSix.SolidDna
{

    public enum CustomPropertyGetResult
    {
        // Cached value was returned
        CachedValue = swCustomInfoGetResult_e.swCustomInfoGetResult_CachedValue,
        // Custom property does not exist
        NotPresent = swCustomInfoGetResult_e.swCustomInfoGetResult_NotPresent,
        // Resolved value was returned
        ResolvedValue = swCustomInfoGetResult_e.swCustomInfoGetResult_ResolvedValue
    }

    public enum CustomPropertySetResult
    {
        // Success
        OK = swCustomInfoSetResult_e.swCustomInfoSetResult_OK,
        // Custom property does not exist
        NotPresent = swCustomInfoSetResult_e.swCustomInfoSetResult_NotPresent,
        // Specified value has an incorrect typeSpecified value has an incorrect type
        TypeMismatch = swCustomInfoSetResult_e.swCustomInfoSetResult_TypeMismatch,
        // ?? not described in SOLIDWORKS API see: https://help.solidworks.com/2020/English/api/swconst/SOLIDWORKS.Interop.swconst~SOLIDWORKS.Interop.swconst.swCustomInfoSetResult_e.html?verRedirect=1
        // but this is returned if the property is linked for e.g. to parent.
        LinkedProp = swCustomInfoSetResult_e.swCustomInfoSetResult_LinkedProp
    }

    public enum CustomPropertyAddResult
    {
        // Success
        AddedOrChanged = swCustomInfoAddResult_e.swCustomInfoAddResult_AddedOrChanged,
        // Failed to add the custom property
        GenericFail = swCustomInfoAddResult_e.swCustomInfoAddResult_GenericFail,
        // Existing custom property with the same name has a different type
        MismatchAgainstExistingType = swCustomInfoAddResult_e.swCustomInfoAddResult_MismatchAgainstExistingType,
        //Specified value of the custom property does not match the specified type
        MismatchAgainstSpecifiedType = swCustomInfoAddResult_e.swCustomInfoAddResult_MismatchAgainstSpecifiedType,
    }

    public enum CustomPropertyDeleteResult
    {
        // Success
        OK = swCustomInfoDeleteResult_e.swCustomInfoDeleteResult_OK,
        // Custom property does not exist
        NotPresent = swCustomInfoDeleteResult_e.swCustomInfoDeleteResult_NotPresent,
        // ?? not described in SOLIDWORKS API see: http://help.solidworks.com/2020/english/api/swconst/SOLIDWORKS.Interop.swconst~SOLIDWORKS.Interop.swconst.swCustomInfoDeleteResult_e.html?verRedirect=1
        // but this is returned if the property is linked for e.g. to parent.
        LinkedProp = swCustomInfoDeleteResult_e.swCustomInfoDeleteResult_LinkedProp
    }

}
