using SolidWorks.Interop.swconst;

namespace AngelSix.SolidDna
{
    public enum CustomPropertyTypes
    {
        Unknown = swCustomInfoType_e.swCustomInfoUnknown,
        Number = swCustomInfoType_e.swCustomInfoNumber,
        Double = swCustomInfoType_e.swCustomInfoDouble,
        YesOrNo = swCustomInfoType_e.swCustomInfoYesOrNo,
        Text = swCustomInfoType_e.swCustomInfoText,
        Date = swCustomInfoType_e.swCustomInfoDate
    }
}
