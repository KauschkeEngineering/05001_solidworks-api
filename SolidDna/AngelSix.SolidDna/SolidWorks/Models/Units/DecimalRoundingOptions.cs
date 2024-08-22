using SolidWorks.Interop.swconst;

namespace AngelSix.SolidDna
{
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
