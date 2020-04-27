using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System;
using System.Collections.Generic;

namespace AngelSix.SolidDna
{
    /// <summary>
    /// Exposes all Assembly Document calls from a <see cref="Model"/>
    /// </summary>
    public class AssemblyDocument
    {
        public enum ComponentResolveStatus
        {
            ResolveAbortedByUser = swComponentResolveStatus_e.swResolveAbortedByUser,
            ResolveError = swComponentResolveStatus_e.swResolveError,
            ResolveNotPerformed = swComponentResolveStatus_e.swResolveNotPerformed,
            ResolveOk = swComponentResolveStatus_e.swResolveOk
        }

        #region Constants

        public const string FILE_EXTENSION = ".sldasm";

        #endregion

        #region Protected Members

        /// <summary>
        /// The base model document. Note we do not dispose of this (the parent Model will)
        /// </summary>
        protected AssemblyDoc mBaseObject;

        #endregion

        #region Public Properties

        /// <summary>
        /// The raw underlying COM object
        /// WARNING: Use with caution. You must handle all disposal from this point on
        /// </summary>
        public AssemblyDoc UnsafeObject => mBaseObject;

        public int LightWeightComponentCount => mBaseObject.GetLightWeightComponentCount();

        #endregion


        #region Constructor

        /// <summary>
        /// Default constructor
        /// </summary>
        public AssemblyDocument(AssemblyDoc model)
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
                SolidDnaErrorCode.SolidWorksModelAssemblyGetFeatureByNameError,
                Localization.GetString(nameof(SolidDnaErrorCode.SolidWorksModelAssemblyGetFeatureByNameError)));
        }

        public ComponentResolveStatus ResolveAllLightWeightComponents()
        {
            return (ComponentResolveStatus)mBaseObject.ResolveAllLightWeightComponents(false);
        }

        public bool ResolveAllLightWeightChildComponents()
        {
            return mBaseObject.ResolveAllLightweight();
        }

        #endregion

    }
}
