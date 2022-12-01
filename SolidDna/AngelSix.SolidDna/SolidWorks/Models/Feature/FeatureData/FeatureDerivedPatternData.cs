using SolidWorks.Interop.sldworks;
using System.Linq;

namespace AngelSix.SolidDna
{
    /// <summary>
    /// Represents a SolidWorks Derived Pattern feature data
    /// </summary>
    public class FeatureDerivedPatternData : SolidDnaObject<IDerivedPatternFeatureData>
    {
        public ModelFeature PatternFeature { get; set; }
        public ModelFeature[] SeedComponentFeatures { protected set; get; }
        public int SkippedItemCount => BaseObject.GetSkippedItemCount();

        #region Constructor

        /// <summary>
        /// Default constructor
        /// </summary>
        public FeatureDerivedPatternData(object model) : base((IDerivedPatternFeatureData)model)
        {
            var llpSeedComps = (object[])BaseObject.SeedComponentArray;
            SeedComponentFeatures = new ModelFeature[llpSeedComps.Count()];

            for (int i = 0; i < llpSeedComps.Count(); i++)
            {
                SeedComponentFeatures[i] = new ModelFeature((Feature)llpSeedComps[i]);
            }
        }

        #endregion

        public bool LoadPatternFeature(Model TopLevelModel)
        {
            if (AccessSelections(TopLevelModel))
            {
                PatternFeature = new ModelFeature(BaseObject.PatternFeature);
                ReleaseSelectionAccess();
                return true;
            }
            return false;
        }

        private bool AccessSelections(Model TopLevelModel, Component componentToModify = null)
        {
            var status = BaseObject.AccessSelections(TopLevelModel.AsModelDoc2, componentToModify);
            return status;
        }

        private void ReleaseSelectionAccess()
        {
            BaseObject.ReleaseSelectionAccess();
        }

    }
}
