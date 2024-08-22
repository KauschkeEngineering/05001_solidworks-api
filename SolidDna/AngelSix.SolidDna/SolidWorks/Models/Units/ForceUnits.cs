using SolidWorks.Interop.swconst;

namespace AngelSix.SolidDna
{
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
}
