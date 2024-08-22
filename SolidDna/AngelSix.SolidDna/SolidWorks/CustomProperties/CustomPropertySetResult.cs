using SolidWorks.Interop.swconst;

namespace AngelSix.SolidDna
{
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
}
