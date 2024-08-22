using SolidWorks.Interop.swconst;

namespace AngelSix.SolidDna
{
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
