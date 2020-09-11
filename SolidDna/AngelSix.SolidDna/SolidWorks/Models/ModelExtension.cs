using AngelSix.SolidDna.SolidWorks.Models.Units;
using AngelSix.SolidDna.SolidWorks.Models.Units.BasicUnits;
using AngelSix.SolidDna.SolidWorks.Models.Units.MassSectionUnits;
using AngelSix.SolidDna.SolidWorks.Models.Units.MotionUnits;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System;

namespace AngelSix.SolidDna
{
    /// <summary>
    /// Represents a SolidWorks model extension of any type (Drawing, Part or Assembly)
    /// </summary>
    public class ModelExtension : SolidDnaObject<ModelDocExtension>
    {
        #region Public Properties

        /// <summary>
        /// The parent Model for this extension
        /// </summary>
        public Model Parent { get; set; }

        public int NeedRebuild => BaseObject.NeedsRebuild2;

        #endregion

        #region Constructor

        /// <summary>
        /// Default constructor
        /// </summary>
        public ModelExtension(ModelDocExtension model, Model parent) : base(model)
        {
            Parent = parent;
        }

        #endregion

        #region Custom Properties

        /// <summary>
        /// Gets a configuration-specific custom property editor for the specified configuration
        /// If no configuration is specified the default custom property manager is returned
        /// 
        /// NOTE: Custom Property Editor must be disposed of once finished
        /// </summary>
        /// <param name="configuration"></param>
        /// <returns></returns>
        public CustomPropertyEditor CustomPropertyEditor(string configuration = null)
        {
            // TODO: Add error checking and exception catching

            return new CustomPropertyEditor(BaseObject.CustomPropertyManager[configuration]);
        }

        #endregion

        #region Mass

        /// <summary>
        /// Gets the mass properties of a part/assembly
        /// </summary>
        /// <param name="doNotThrowOnError">If true, don't throw on errors, just return empty mass</param>
        /// <returns></returns>
        public MassProperties GetMassProperties(bool doNotThrowOnError = true)
        {
            // Wrap any error
            return SolidDnaErrors.Wrap(() =>
            {
                // Make sure we are a part
                if (!Parent.IsPart && !Parent.IsAssembly)
                {
                    if (doNotThrowOnError)
                        return new MassProperties();
                    else
                        throw new InvalidOperationException(Localization.GetString("SolidWorksModelGetMassModelNotPartError"));
                }

                double[] massProps = null;
                var status = -1;

                //
                // SolidWorks 2016 is the start of support for MassProperties2
                //
                // Tested on 2015 and crashes, so drop-back to lower version for support
                //
                if (SolidWorksEnvironment.Application.SolidWorksVersion.Version < 2016)
                    // NOTE: 2 is best accuracy
                    massProps = (double[])BaseObject.GetMassProperties(2, ref status);
                else
                    // NOTE: 2 is best accuracy
                    massProps = (double[])BaseObject.GetMassProperties2(2, out status, false);

                // Make sure it succeeded
                if (status == (int)swMassPropertiesStatus_e.swMassPropertiesStatus_UnknownError)
                {
                    if (doNotThrowOnError)
                        return new MassProperties();
                    else
                        throw new InvalidOperationException(Localization.GetString("SolidWorksModelGetMassModelStatusFailed"));
                }
                // If we have no mass, return empty
                else if (status == (int)swMassPropertiesStatus_e.swMassPropertiesStatus_NoBody)
                    return new MassProperties();

                // Otherwise we have the properties so return them
                return new MassProperties(massProps);
            },
                SolidDnaErrorTypeCode.SolidWorksModel,
                SolidDnaErrorCode.SolidWorksModelGetMassPropertiesError,
                Localization.GetString("SolidWorksModelGetMassPropertiesError"));
        }

        #endregion

        #region Dispose

        public override void Dispose()
        {
            // Clear reference to be safe
            Parent = null;

            base.Dispose();
        }

        #endregion

        public bool Rebuild(swRebuildOptions_e buildOption)
        {
            return ((ModelDocExtension)mBaseObject).Rebuild((int)buildOption);
        }

        public bool Rename(string oldName, string newName)
        {
            if (BaseObject != null)
            {
                if (BaseObject.SelectByID2(oldName, "COMPONENT", 0, 0, 0, false, 0, null, 0))
                    if (BaseObject.RenameDocument(newName) == 0)
                        return true;
            }
            return false;

        }

        public bool Select()
        {
            return BaseObject != null ? BaseObject.SelectByID2(Parent.Name, "COMPONENT", 0, 0, 0, false, 0, null, 0) : false;
        }

        public void SetIsometricZoomToFitView(bool showIsometric)
        {
            if (BaseObject != null)
            {
                BaseObject.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swImageQualityZoomToFitForPreviewImages, (int)swUserPreferenceOption_e.swDetailingNoOptionSpecified, showIsometric);
            }
        }

        #region Get/Set Unit System

        public UnitSystems GetUnitSystem()
        {
            if (BaseObject == null)
                return UnitSystems.NotAvailable;
            return (UnitSystems)BaseObject.GetUserPreferenceInteger((int)swUserPreferenceIntegerValue_e.swUnitSystem, (int)swUserPreferenceOption_e.swDetailingNoOptionSpecified);
        }

        public bool SetUnitSystem(UnitSystems unitSystem)
        {
            if (BaseObject == null)
                return false;
            return BaseObject.SetUserPreferenceInteger((int)swUserPreferenceIntegerValue_e.swUnitSystem, (int)swUserPreferenceOption_e.swDetailingNoOptionSpecified, (int)unitSystem);
        }

        #endregion

        #region Get/Set Length Unit and Decimal Places

        public LengthUnits GetLengthUnit()
        {
            if (BaseObject == null)
                return LengthUnits.NotAvailable;
            return (LengthUnits)BaseObject.GetUserPreferenceInteger((int)swUserPreferenceIntegerValue_e.swUnitsLinear, (int)swUserPreferenceOption_e.swDetailingNoOptionSpecified);
        }

        public bool SetLengthUnit(LengthUnits lengthUnit)
        {
            if (BaseObject == null)
                return false;
            return BaseObject.SetUserPreferenceInteger((int)swUserPreferenceIntegerValue_e.swUnitsLinear, (int)swUserPreferenceOption_e.swDetailingNoOptionSpecified, (int)lengthUnit);
        }

        public int GetLengthDecimalPlaces()
        {
            if (BaseObject == null)
                return -1;
            return BaseObject.GetUserPreferenceInteger((int)swUserPreferenceIntegerValue_e.swUnitsLinearDecimalPlaces, (int)swUserPreferenceOption_e.swDetailingNoOptionSpecified);
        }

        public bool SetLengthDecimalPlaces(int decimalPlaces)
        {
            if (BaseObject == null)
                return false;
            return BaseObject.SetUserPreferenceInteger((int)swUserPreferenceIntegerValue_e.swUnitsLinearDecimalPlaces, (int)swUserPreferenceOption_e.swDetailingNoOptionSpecified, decimalPlaces);
        }

        #endregion

        #region Get/Set Dual Dimensional Length Unit and Decimal Places

        public LengthUnits GetDualDimensionalLengthUnit()
        {
            if (BaseObject == null)
                return LengthUnits.NotAvailable;
            return (LengthUnits)BaseObject.GetUserPreferenceInteger((int)swUserPreferenceIntegerValue_e.swUnitsDualLinear, (int)swUserPreferenceOption_e.swDetailingNoOptionSpecified);
        }

        public bool SetDualDimensionalLengthUnit(LengthUnits lengthUnit)
        {
            if (BaseObject == null)
                return false;
            return BaseObject.SetUserPreferenceInteger((int)swUserPreferenceIntegerValue_e.swUnitsDualLinear, (int)swUserPreferenceOption_e.swDetailingNoOptionSpecified, (int)lengthUnit);
        }

        public int GetDualDimensionalLengthDecimalPlaces()
        {
            if (BaseObject == null)
                return -1;
            return BaseObject.GetUserPreferenceInteger((int)swUserPreferenceIntegerValue_e.swUnitsDualLinearDecimalPlaces, (int)swUserPreferenceOption_e.swDetailingNoOptionSpecified);
        }

        public bool SetDualDimensionalLengthDecimalPlaces(int decimalPlaces)
        {
            if (BaseObject == null)
                return false;
            return BaseObject.SetUserPreferenceInteger((int)swUserPreferenceIntegerValue_e.swUnitsDualLinearDecimalPlaces, (int)swUserPreferenceOption_e.swDetailingNoOptionSpecified, decimalPlaces);
        }

        #endregion

        #region Get/Set Angle Unit and Decimal Places

        public AngleUnits GetAngleUnit()
        {
            if (BaseObject == null)
                return AngleUnits.NotAvailable;
            return (AngleUnits)BaseObject.GetUserPreferenceInteger((int)swUserPreferenceIntegerValue_e.swUnitsAngular, (int)swUserPreferenceOption_e.swDetailingNoOptionSpecified);
        }

        public bool SetAngleUnit(AngleUnits angleUnit)
        {
            if (BaseObject == null)
                return false;
            return BaseObject.SetUserPreferenceInteger((int)swUserPreferenceIntegerValue_e.swUnitsAngular, (int)swUserPreferenceOption_e.swDetailingNoOptionSpecified, (int)angleUnit);
        }

        public int GetAngleDecimalPlaces()
        {
            if (BaseObject == null)
                return -1;
            return BaseObject.GetUserPreferenceInteger((int)swUserPreferenceIntegerValue_e.swUnitsAngularDecimalPlaces, (int)swUserPreferenceOption_e.swDetailingNoOptionSpecified);
        }

        public bool SetAngleDecimalPlaces(int decimalPlaces)
        {
            if (BaseObject == null)
                return false;
            return BaseObject.SetUserPreferenceInteger((int)swUserPreferenceIntegerValue_e.swUnitsAngularDecimalPlaces, (int)swUserPreferenceOption_e.swDetailingNoOptionSpecified, decimalPlaces);
        }

        #endregion

        #region Get/Set Mass Section Length Unit and Decimal Places

        public LengthUnits GetMassSectionLengthUnit()
        {
            if (BaseObject == null)
                return LengthUnits.NotAvailable;
            return (LengthUnits)BaseObject.GetUserPreferenceInteger((int)swUserPreferenceIntegerValue_e.swUnitsMassPropLength, (int)swUserPreferenceOption_e.swDetailingNoOptionSpecified);
        }

        public bool SetMassSectionLengthUnit(LengthUnits lengthUnit)
        {
            if (BaseObject == null)
                return false;
            return BaseObject.SetUserPreferenceInteger((int)swUserPreferenceIntegerValue_e.swUnitsMassPropLength, (int)swUserPreferenceOption_e.swDetailingNoOptionSpecified, (int)lengthUnit);
        }

        public int GetMassSectionLengthDecimalPlaces()
        {
            if (BaseObject == null)
                return -1;
            return BaseObject.GetUserPreferenceInteger((int)swUserPreferenceIntegerValue_e.swUnitsMassPropDecimalPlaces, (int)swUserPreferenceOption_e.swDetailingNoOptionSpecified);
        }

        public bool SetMassSectionLengthDecimalPlaces(int decimalPlaces)
        {
            if (BaseObject == null)
                return false;
            return BaseObject.SetUserPreferenceInteger((int)swUserPreferenceIntegerValue_e.swUnitsMassPropDecimalPlaces, (int)swUserPreferenceOption_e.swDetailingNoOptionSpecified, decimalPlaces);
        }

        #endregion

        #region Get/Set Mass Section Mass Unit

        public MassUnits GetMassSectionMassUnit()
        {
            if (BaseObject == null)
                return MassUnits.NotAvailable;
            return (MassUnits)BaseObject.GetUserPreferenceInteger((int)swUserPreferenceIntegerValue_e.swUnitsMassPropMass, (int)swUserPreferenceOption_e.swDetailingNoOptionSpecified);
        }

        public bool SetMassSectionMassUnit(MassUnits massUnit)
        {
            if (BaseObject == null)
                return false;
            return BaseObject.SetUserPreferenceInteger((int)swUserPreferenceIntegerValue_e.swUnitsMassPropMass, (int)swUserPreferenceOption_e.swDetailingNoOptionSpecified, (int)massUnit);
        }

        #endregion

        #region Get/Set Mass Section Mass Unit

        public MassVolumeUnits GetMassSectionVolumeUnit()
        {
            if (BaseObject == null)
                return MassVolumeUnits.NotAvailable;
            return (MassVolumeUnits)BaseObject.GetUserPreferenceInteger((int)swUserPreferenceIntegerValue_e.swUnitsMassPropVolume, (int)swUserPreferenceOption_e.swDetailingNoOptionSpecified);
        }

        public bool SetMassSectionVolumeUnit(MassVolumeUnits massVolumeUnit)
        {
            if (BaseObject == null)
                return false;
            return BaseObject.SetUserPreferenceInteger((int)swUserPreferenceIntegerValue_e.swUnitsMassPropVolume, (int)swUserPreferenceOption_e.swDetailingNoOptionSpecified, (int)massVolumeUnit);
        }

        #endregion

        #region Get/Set Time Unit and Decimal Places

        public TimeUnits GetTimeUnit()
        {
            if (BaseObject == null)
                return TimeUnits.NotAvailable;
            return (TimeUnits)BaseObject.GetUserPreferenceInteger((int)swUserPreferenceIntegerValue_e.swUnitsTimeUnits, (int)swUserPreferenceOption_e.swDetailingNoOptionSpecified);
        }

        public bool SetTimeUnit(TimeUnits timeUnit)
        {
            if (BaseObject == null)
                return false;
            return BaseObject.SetUserPreferenceInteger((int)swUserPreferenceIntegerValue_e.swUnitsTimeUnits, (int)swUserPreferenceOption_e.swDetailingNoOptionSpecified, (int)timeUnit);
        }

        public int GetTimeDecimalPlaces()
        {
            if (BaseObject == null)
                return -1;
            return BaseObject.GetUserPreferenceInteger((int)swUserPreferenceIntegerValue_e.swUnitsTimeDecimalPlaces, (int)swUserPreferenceOption_e.swDetailingNoOptionSpecified);
        }

        public bool SetTimeDecimalPlaces(int decimalPlaces)
        {
            if (BaseObject == null)
                return false;
            return BaseObject.SetUserPreferenceInteger((int)swUserPreferenceIntegerValue_e.swUnitsTimeDecimalPlaces, (int)swUserPreferenceOption_e.swDetailingNoOptionSpecified, decimalPlaces);
        }

        #endregion

        #region Get/Set Force Unit and Decimal Places

        public ForceUnits GetForceUnit()
        {
            if (BaseObject == null)
                return ForceUnits.NotAvailable;
            return (ForceUnits)BaseObject.GetUserPreferenceInteger((int)swUserPreferenceIntegerValue_e.swUnitsForce, (int)swUserPreferenceOption_e.swDetailingNoOptionSpecified);
        }

        public bool SetForceUnit(ForceUnits forceUnit)
        {
            if (BaseObject == null)
                return false;
            return BaseObject.SetUserPreferenceInteger((int)swUserPreferenceIntegerValue_e.swUnitsForce, (int)swUserPreferenceOption_e.swDetailingNoOptionSpecified, (int)forceUnit);
        }

        public int GetForceDecimalPlaces()
        {
            if (BaseObject == null)
                return -1;
            return BaseObject.GetUserPreferenceInteger((int)swUserPreferenceIntegerValue_e.swUnitsForceDecimalPlaces, (int)swUserPreferenceOption_e.swDetailingNoOptionSpecified);
        }

        public bool SetForceDecimalPlaces(int decimalPlaces)
        {
            if (BaseObject == null)
                return false;
            return BaseObject.SetUserPreferenceInteger((int)swUserPreferenceIntegerValue_e.swUnitsForceDecimalPlaces, (int)swUserPreferenceOption_e.swDetailingNoOptionSpecified, decimalPlaces);
        }

        #endregion

        #region Get/Set Power Unit and Decimal Places

        public PowerUnits GetPowerUnit()
        {
            if (BaseObject == null)
                return PowerUnits.NotAvailable;
            return (PowerUnits)BaseObject.GetUserPreferenceInteger((int)swUserPreferenceIntegerValue_e.swUnitsPowerUnits, (int)swUserPreferenceOption_e.swDetailingNoOptionSpecified);
        }

        public bool SetPowerUnit(PowerUnits powerUnits)
        {
            if (BaseObject == null)
                return false;
            return BaseObject.SetUserPreferenceInteger((int)swUserPreferenceIntegerValue_e.swUnitsPowerUnits, (int)swUserPreferenceOption_e.swDetailingNoOptionSpecified, (int)powerUnits);
        }

        public int GetPowerDecimalPlaces()
        {
            if (BaseObject == null)
                return -1;
            return BaseObject.GetUserPreferenceInteger((int)swUserPreferenceIntegerValue_e.swUnitsPowerDecimalPlaces, (int)swUserPreferenceOption_e.swDetailingNoOptionSpecified);
        }

        public bool SetPowerDecimalPlaces(int decimalPlaces)
        {
            if (BaseObject == null)
                return false;
            return BaseObject.SetUserPreferenceInteger((int)swUserPreferenceIntegerValue_e.swUnitsPowerDecimalPlaces, (int)swUserPreferenceOption_e.swDetailingNoOptionSpecified, decimalPlaces);
        }

        #endregion

        #region Get/Set Energy Unit and Decimal Places

        public EnergyUnits GetEnergyUnit()
        {
            if (BaseObject == null)
                return EnergyUnits.NotAvailable;
            return (EnergyUnits)BaseObject.GetUserPreferenceInteger((int)swUserPreferenceIntegerValue_e.swUnitsEnergyUnits, (int)swUserPreferenceOption_e.swDetailingNoOptionSpecified);
        }

        public bool SetEnergyUnit(EnergyUnits energyUnit)
        {
            if (BaseObject == null)
                return false;
            return BaseObject.SetUserPreferenceInteger((int)swUserPreferenceIntegerValue_e.swUnitsEnergyUnits, (int)swUserPreferenceOption_e.swDetailingNoOptionSpecified, (int)energyUnit);
        }

        public int GetEnergyDecimalPlaces()
        {
            if (BaseObject == null)
                return -1;
            return BaseObject.GetUserPreferenceInteger((int)swUserPreferenceIntegerValue_e.swUnitsEnergyDecimalPlaces, (int)swUserPreferenceOption_e.swDetailingNoOptionSpecified);
        }

        public bool SetEnergyDecimalPlaces(int decimalPlaces)
        {
            if (BaseObject == null)
                return false;
            return BaseObject.SetUserPreferenceInteger((int)swUserPreferenceIntegerValue_e.swUnitsEnergyDecimalPlaces, (int)swUserPreferenceOption_e.swDetailingNoOptionSpecified, decimalPlaces);
        }

        #endregion

    }
}
