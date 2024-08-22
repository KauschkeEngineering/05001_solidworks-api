using SolidWorks.Interop.swconst;

namespace AngelSix.SolidDna
{
    /// <summary>
    /// Custom property types.
    /// <see cref="swCustomInfoType_e"/>
    /// </summary>
    public enum CustomPropertyTypes
    {
        // type is not known
        Unknown = swCustomInfoType_e.swCustomInfoUnknown,
        // type is of number
        Number = swCustomInfoType_e.swCustomInfoNumber,
        // type is of double
        Double = swCustomInfoType_e.swCustomInfoDouble,
        // type is of yes no
        YesOrNo = swCustomInfoType_e.swCustomInfoYesOrNo,
        // type is of text
        Text = swCustomInfoType_e.swCustomInfoText,
        // type is of data and time
        Date = swCustomInfoType_e.swCustomInfoDate
    }
}
