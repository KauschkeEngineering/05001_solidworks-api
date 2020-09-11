using SolidWorks.Interop.swconst;

namespace AngelSix.SolidDna.SolidWorks.Models.Units.BasicUnits
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

    /// <summary>
    /// Specifies decimal or fraction display for linear units
    /// <see cref="swFractionDisplay_e "/>
    /// </summary>
    public enum FractionDisplay
    {
        NotAvailable = -1,
        None = swFractionDisplay_e.swNONE,
        Decimal= swFractionDisplay_e.swDECIMAL,
        Fraction = swFractionDisplay_e.swFRACTION
    }

    /// <summary>
    /// Specifies angular units
    /// <see cref="swLengthUnit_e"/>
    /// </summary>
    public enum AngleUnits
    {
        NotAvailable = -1,
        Degrees = swAngleUnit_e.swDEGREES,
        DegreeMinutes = swAngleUnit_e.swDEG_MIN,
        DegreeMinutesSeconds = swAngleUnit_e.swDEG_MIN_SEC,
        Radians = swAngleUnit_e.swRADIANS
    }
}
