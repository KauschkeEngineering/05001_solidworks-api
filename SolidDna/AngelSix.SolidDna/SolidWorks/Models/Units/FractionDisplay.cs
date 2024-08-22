using SolidWorks.Interop.swconst;

namespace AngelSix.SolidDna
{
    /// <summary>
    /// Specifies decimal or fraction display for linear units
    /// <see cref="swFractionDisplay_e "/>
    /// </summary>
    public enum FractionDisplay
    {
        NotAvailable = -1,
        None = swFractionDisplay_e.swNONE,
        Decimal = swFractionDisplay_e.swDECIMAL,
        Fraction = swFractionDisplay_e.swFRACTION
    }
}
