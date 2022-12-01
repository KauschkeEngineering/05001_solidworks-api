using SolidWorks.Interop.sldworks;
using System.Linq;

namespace AngelSix.SolidDna
{
    /// <summary>
    /// Represents a SolidWorks Circular Pattern feature data
    /// </summary>
    public class FeatureCircularPatternData : SolidDnaObject<ICircularPatternFeatureData>
    {
        public int SeedComponentCount => BaseObject.GetPatternFeatureCount();
        public int SkippedItemCount => BaseObject.GetSkippedItemCount();
        public int D1TotalInstances => BaseObject.TotalInstances;
        public int D2TotalInstances => BaseObject.TotalInstances2;
        public bool Direction2Enabled => BaseObject.Direction2;
        public ModelFeature[] PatternComponentFeatures { protected set; get; }

        #region Constructor

        /// <summary>
        /// Default constructor
        /// </summary>
        public FeatureCircularPatternData(object model) : base((ICircularPatternFeatureData)model)
        {
            var llpSeedComps = (object[])BaseObject.PatternFeatureArray;
            PatternComponentFeatures = new ModelFeature[llpSeedComps.Count()];

            for (int i = 0; i < llpSeedComps.Count(); i++)
            {
                PatternComponentFeatures[i] = new ModelFeature((Feature)llpSeedComps[i]);
            }
        }

        #endregion
    }
}
