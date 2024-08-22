using SolidWorks.Interop.swconst;

namespace AngelSix.SolidDna
{
    /// <summary>
    /// Specifies type of system unit
    /// <see cref="swUnitSystem_e"/>
    /// </summary>
    public enum UnitSystems
    {
        NotAvailable = -1,
        // centimeter, gram, second
        CentimeterGramSecond = swUnitSystem_e.swUnitSystem_CGS,
        // meter, kilogram, second
        MeterKilogramSecond = swUnitSystem_e.swUnitSystem_MKS,
        // inch, pound, second
        InchPoundSecond = swUnitSystem_e.swUnitSystem_IPS,
        // lets you set length units, density units, and force
        Custom = swUnitSystem_e.swUnitSystem_Custom,
        // millimeter, gram, second
        MillimeterGramSecond = swUnitSystem_e.swUnitSystem_MMGS
    }
}
