using SolidWorks.Interop.sldworks;
using System.Linq;

namespace AngelSix.SolidDna
{
    /// <summary>
    /// Represents a SolidWorks Mirror Pattern feature data
    /// </summary>
    public class FeatureMirrorPatternData : SolidDnaObject<IMirrorPatternFeatureData>
    {
        public ModelFeature[] PatternComponentFeatures { protected set; get; }

        #region Constructor

        /// <summary>
        /// Default constructor
        /// </summary>
        public FeatureMirrorPatternData(object model) : base((IMirrorPatternFeatureData)model)
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
