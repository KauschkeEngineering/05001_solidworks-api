using SolidWorks.Interop.swconst;

namespace AngelSix.SolidDna
{
    /// <summary>
    /// Specifies the type of time unit.
    /// NOTE: You can only set this option when swUnitSystem_e is swUnitSytem_Custom;
    /// Otherwise, FALSE is returned
    /// <see cref="swUnitsTimeUnit_e"/>
    /// </summary>
    public enum TimeUnits
    {
        NotAvailable = -1,
        Second = swUnitsTimeUnit_e.swUnitsTimeUnit_Second,
        Millisecond = swUnitsTimeUnit_e.swUnitsTimeUnit_Millisecond,
        Minute = swUnitsTimeUnit_e.swUnitsTimeUnit_Minute,
        Hour = swUnitsTimeUnit_e.swUnitsTimeUnit_Hour,
        Microsecond = swUnitsTimeUnit_e.swUnitsTimeUnit_Microsecond,
        Nanosecond = swUnitsTimeUnit_e.swUnitsTimeUnit_Nanosecond
    }
}
