using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System;
using System.Collections.Generic;
using System.Linq;

namespace AngelSix.SolidDna
{
    /// <summary>
    /// Exposes all Part Document calls from a <see cref="Model"/>
    /// </summary>
    public class PartDocument
    {

        #region Constants

        public const string FILE_EXTENSION = ".sldprt";

        #endregion

        #region Protected Members

        /// <summary>
        /// The base model document. Note we do not dispose of this (the parent Model will)
        /// </summary>
        protected PartDoc mBaseObject;

        #endregion

        #region Public Properties

        /// <summary>
        /// The raw underlying COM object
        /// WARNING: Use with caution. You must handle all disposal from this point on
        /// </summary>
        public PartDoc UnsafeObject => mBaseObject;

        #endregion

        #region Constructor

        /// <summary>
        /// Default constructor
        /// </summary>
        public PartDocument(PartDoc model)
        {
            mBaseObject = model;
        }

        #endregion

        #region Feature Methods

        /// <summary>
        /// Gets the <see cref="ModelFeature"/> of the item in the feature tree based on its name
        /// </summary>
        /// <param name="featureName">Name of the feature</param>
        /// <returns>The <see cref="ModelFeature"/> for the named feature</returns>
        public void GetFeatureByName(string featureName, Action<ModelFeature> action)
        {
            // Wrap any error
            SolidDnaErrors.Wrap(() =>
            {
                // Create feature
                using (var model = new ModelFeature((Feature)mBaseObject.FeatureByName(featureName)))
                {
                    // Run action
                    action(model);
                }
            },
                SolidDnaErrorTypeCode.SolidWorksModel,
                SolidDnaErrorCode.SolidWorksModelPartGetFeatureByNameError,
                Localization.GetString(nameof(SolidDnaErrorCode.SolidWorksModelPartGetFeatureByNameError)));
        }

        // TODO: add error code for new method

        public Material GetMaterial(List<MaterialDatabase> materialDatabases, string configName = "")
        {
            // Wrap any error
            return SolidDnaErrors.Wrap(() =>
            {
                string usedDatabase = "";
                string usedMaterialName = mBaseObject.GetMaterialPropertyName2(configName, out usedDatabase);
                //MaterialVisualPropertiesData myMatVisProps = mBaseObject.GetMaterialVisualProperties();
                var id = mBaseObject.MaterialIdName;

                if (usedDatabase.Equals("") && usedMaterialName.Equals(""))
                    return new Material();
                else
                {
                    var materialDatabasePath = materialDatabases.FirstOrDefault(database => database.Name.ToLower().Equals(usedDatabase.ToLower()));
                    return materialDatabasePath == null
                        ? new Material() { Database = new MaterialDatabase(usedDatabase), Name = usedMaterialName }
                        : AddInIntegration.SolidWorks.FindMaterial(materialDatabases.FirstOrDefault(database => database.Name.ToLower().Equals(usedDatabase.ToLower())).FullPath, usedMaterialName);
                }
            },
                SolidDnaErrorTypeCode.SolidWorksModel,
                SolidDnaErrorCode.SolidWorksModelPartGetFeatureByNameError,
                Localization.GetString(nameof(SolidDnaErrorCode.SolidWorksModelPartGetFeatureByNameError)));
        }

        // TODO: add error code for new method

        public void SetMaterialDatabase(string databaseName, string configName = "")
        {
            // Wrap any error
            SolidDnaErrors.Wrap(() =>
            {
                mBaseObject.SetMaterialPropertyName2(configName, databaseName, "");
            },
                SolidDnaErrorTypeCode.SolidWorksModel,
                SolidDnaErrorCode.SolidWorksModelPartGetFeatureByNameError,
                Localization.GetString(nameof(SolidDnaErrorCode.SolidWorksModelPartGetFeatureByNameError)));
        }

        public void SetMaterial(string databaseName, string materialName, string configName = "")
        {
            // Wrap any error
            SolidDnaErrors.Wrap(() =>
            {
                mBaseObject.SetMaterialPropertyName2(configName, databaseName, materialName);
            },
                SolidDnaErrorTypeCode.SolidWorksModel,
                SolidDnaErrorCode.SolidWorksModelPartGetFeatureByNameError,
                Localization.GetString(nameof(SolidDnaErrorCode.SolidWorksModelPartGetFeatureByNameError)));
        }

        #endregion
    }
}
