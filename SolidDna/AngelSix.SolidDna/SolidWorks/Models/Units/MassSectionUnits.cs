using SolidWorks.Interop.swconst;

namespace AngelSix.SolidDna.SolidWorks.Models.Units.MassSectionUnits
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

    /// <summary>
    /// Specifies type of volume unit to use for mass property units
    /// NOTE: You can only set this option when swUnitSystem_e is swUnitSytem_Custom;
    /// Otherwise, FALSE is returned
    /// <see cref="swUnitsMassPropMass_e"/>
    /// </summary>
    public enum MassVolumeUnits
    {
        NotAvailable = -1,
        Angstroms3 = swUnitsMassPropVolume_e.swUnitsMassPropVolume_Angstroms3,
        Nanometers3 = swUnitsMassPropVolume_e.swUnitsMassPropVolume_Nanometers3,
        Microns3 = swUnitsMassPropVolume_e.swUnitsMassPropVolume_Microns3,
        Millimeters3 = swUnitsMassPropVolume_e.swUnitsMassPropVolume_Millimeters3,
        Centimeters3 = swUnitsMassPropVolume_e.swUnitsMassPropVolume_Centimeters3,
        Meters3 = swUnitsMassPropVolume_e.swUnitsMassPropVolume_Meters3,
        Microinches3 = swUnitsMassPropVolume_e.swUnitsMassPropVolume_Microinches3,
        Mils3 = swUnitsMassPropVolume_e.swUnitsMassPropVolume_Mils3,
        Inches3 = swUnitsMassPropVolume_e.swUnitsMassPropVolume_Inches3,
        Feet3 = swUnitsMassPropVolume_e.swUnitsMassPropVolume_Feet3,
        MicroLiters = swUnitsMassPropVolume_e.swUnitsMassPropVolume_MicroLiters,
        MilliLiters = swUnitsMassPropVolume_e.swUnitsMassPropVolume_MilliLiters,
        CentiLiters = swUnitsMassPropVolume_e.swUnitsMassPropVolume_CentiLiters,
        DeciLiters = swUnitsMassPropVolume_e.swUnitsMassPropVolume_DeciLiters,
        Liters = swUnitsMassPropVolume_e.swUnitsMassPropVolume_Liters,
        HectoLiters = swUnitsMassPropVolume_e.swUnitsMassPropVolume_HectoLiters,
        USFluidOunce = swUnitsMassPropVolume_e.swUnitsMassPropVolume_USFluidOunce,
        USPints = swUnitsMassPropVolume_e.swUnitsMassPropVolume_USPints,
        USGallons = swUnitsMassPropVolume_e.swUnitsMassPropVolume_USGallons,
        IMPGallons = swUnitsMassPropVolume_e.swUnitsMassPropVolume_IMPGallons,
        IMPCubicYards = swUnitsMassPropVolume_e.swUnitsMassPropVolume_IMPCubicYards
    }
}
