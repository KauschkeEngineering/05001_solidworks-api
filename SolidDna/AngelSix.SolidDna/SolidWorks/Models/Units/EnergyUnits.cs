using SolidWorks.Interop.swconst;

namespace AngelSix.SolidDna
{
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
