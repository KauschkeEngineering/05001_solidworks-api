using SolidWorks.Interop.sldworks;
using System.Linq;

namespace AngelSix.SolidDna
{
    /// <summary>
    /// Represents a SolidWorks Local Linear Pattern feature data
    /// </summary>
    public class FeatureLocalLinearPatternData : SolidDnaObject<ILocalLinearPatternFeatureData>
    {
        public int SeedComponentCount => BaseObject.GetSeedComponentCount();
        public int SkippedItemCount => BaseObject.GetSkippedItemCount();
        public int D1TotalInstances => BaseObject.D1TotalInstances;
        public int D2TotalInstances => BaseObject.D2TotalInstances;
        public ModelFeature[] SeedComponentFeatures { protected set; get; }

        #region Constructor

        /// <summary>
        /// Default constructor
        /// </summary>
        public FeatureLocalLinearPatternData(object model) : base((ILocalLinearPatternFeatureData)model)
        {
            var llpSeedComps = (object[])BaseObject.SeedComponentArray;
            SeedComponentFeatures = new ModelFeature[llpSeedComps.Count()];

            for (int i = 0; i < llpSeedComps.Count(); i++)
            {
                SeedComponentFeatures[i] = new ModelFeature((Feature)llpSeedComps[i]);
            }
        }

        #endregion
    }
}
