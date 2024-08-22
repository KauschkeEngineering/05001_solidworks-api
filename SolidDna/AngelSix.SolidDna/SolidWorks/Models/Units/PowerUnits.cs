using SolidWorks.Interop.swconst;

namespace AngelSix.SolidDna
{
    /// <summary>
    /// Specifies type of power unit.
    /// NOTE: You can only set this option when swUnitSystem_e is swUnitSytem_Custom;
    /// Otherwise, FALSE is returned
    /// <see cref="swUnitsPowerUnit_e"/>
    /// </summary>
    public enum PowerUnits
    {
        NotAvailable = -1,
        Watt = swUnitsPowerUnit_e.swUnitsPowerUnit_Watt,
        Horsepower = swUnitsPowerUnit_e.swUnitsPowerUnit_Horsepower,
        Kilowatt = swUnitsPowerUnit_e.swUnitsPowerUnit_Kilowatt
    }
}
