using SolidWorks.Interop.swconst;

namespace AngelSix.SolidDna
{
    /// <summary>
    /// Specifies type of mass unit to use for mass property units
    /// NOTE: You can only set this option when swUnitSystem_e is swUnitSytem_Custom; 
    /// Otherwise, FALSE is returned
    /// <see cref="swUnitsMassPropMass_e"/>
    /// </summary>
    public enum MassUnits
    {
        NotAvailable = -1,
        Milligrams = swUnitsMassPropMass_e.swUnitsMassPropMass_Milligrams,
        Grams = swUnitsMassPropMass_e.swUnitsMassPropMass_Grams,
        Kilograms = swUnitsMassPropMass_e.swUnitsMassPropMass_Kilograms,
        Pounds = swUnitsMassPropMass_e.swUnitsMassPropMass_Pounds
    }
}
