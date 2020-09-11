using SolidWorks.Interop.swconst;

namespace AngelSix.SolidDna.SolidWorks.Models.Units.MotionUnits
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

    /// <summary>
    /// Specifies the type of force units.
    /// NOTE: You can only set this option when swUnitSystem_e is swUnitSytem_Custom; 
    /// Otherwise, FALSE is returned
    /// <see cref="swUnitsForce_e"/>
    /// </summary>
    public enum ForceUnits
    {
        NotAvailable = -1,
        Dynes = swUnitsForce_e.swUnitsForce_Dynes,
        Millinewtons = swUnitsForce_e.swUnitsForce_Millinewtons,
        Newtons = swUnitsForce_e.swUnitsForce_Newtons,
        Kilonewtons = swUnitsForce_e.swUnitsForce_Kilonewtons,
        Meganewtons = swUnitsForce_e.swUnitsForce_Meganewtons,
        Poundfeet = swUnitsForce_e.swUnitsForce_Poundfeet,
        KgForce = swUnitsForce_e.swUnitsForce_KgForce,
        OunceForce = swUnitsForce_e.swUnitsForce_OunceForce
    }

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

    /// <summary>
    /// Specifies the type of energy units.
    /// NOTE: You can only set this option when swUnitSystem_e is swUnitSytem_Custom; 
    /// Otherwise, FALSE is returned
    /// <see cref="swUnitsEnergyUnit_e"/>
    /// </summary>
    public enum EnergyUnits
    {
        NotAvailable = -1,
        Joule = swUnitsEnergyUnit_e.swUnitsEnergyUnit_Joule,
        Erg = swUnitsEnergyUnit_e.swUnitsEnergyUnit_Ergs,
        BritisThermalUnit = swUnitsEnergyUnit_e.swUnitsEnergyUnit_BTU,
        KilowattHour = swUnitsEnergyUnit_e.swUnitsEnergyUnit_KilowattHour
    }

}
