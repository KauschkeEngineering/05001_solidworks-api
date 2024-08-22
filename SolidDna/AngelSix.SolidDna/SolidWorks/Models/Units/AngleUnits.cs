using SolidWorks.Interop.swconst;

namespace AngelSix.SolidDna
{
    /// <summary>
    /// Specifies angular units
    /// <see cref="swLengthUnit_e"/>
    /// </summary>
    public enum AngleUnits
    {
        NotAvailable = -1,
        Degrees = swAngleUnit_e.swDEGREES,
        DegreeMinutes = swAngleUnit_e.swDEG_MIN,
        DegreeMinutesSeconds = swAngleUnit_e.swDEG_MIN_SEC,
        Radians = swAngleUnit_e.swRADIANS
    }
}
