using SolidWorks.Interop.swconst;

namespace AngelSix.SolidDna.SolidWorks.Models.Units
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

    /// <summary>
    /// Specifies decimal rounding method
    /// <see cref="swUnitsDecimalRounding_e"/>
    /// </summary>
    public enum DecimalRoundingOptions
    {
        // round up to the nearest decimal
        HalfAway = swUnitsDecimalRounding_e.swUnitsDecimalRounding_HalfAway,
        // round down to the nearest decimal
        HalfTowards = swUnitsDecimalRounding_e.swUnitsDecimalRounding_HalfTowards,
        // round up or down to the next even decimal
        HalfToEven = swUnitsDecimalRounding_e.swUnitsDecimalRounding_HalfToEven,
        // truncate the decimal without rounding
        Truncate = swUnitsDecimalRounding_e.swUnitsDecimalRounding_Truncate
    }
}
