using SolidWorks.Interop.swconst;

namespace AngelSix.SolidDna
{
    /// <summary>
    /// Specifies the type of linear length units
    /// <see cref="swLengthUnit_e"/>
    /// </summary>
    public enum LengthUnits
    {
        NotAvailable = -1,
        Millimeter = swLengthUnit_e.swMM,
        Centimeter = swLengthUnit_e.swCM,
        Meter = swLengthUnit_e.swMETER,
        Inch = swLengthUnit_e.swINCHES,
        Feet = swLengthUnit_e.swFEET,
        FeetAndInch = swLengthUnit_e.swFEETINCHES,
        Angstrom = swLengthUnit_e.swANGSTROM,
        Nanometer = swLengthUnit_e.swNANOMETER,
        Micrometer = swLengthUnit_e.swMICRON,
        Milliinch = swLengthUnit_e.swMIL,
        Microinch = swLengthUnit_e.swUIN
    }
}
