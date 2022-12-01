using SolidWorks.Interop.sldworks;
using System.Linq;

namespace AngelSix.SolidDna
{
    /// <summary>
    /// Represents a SolidWorks Local Curve Pattern feature data
    /// </summary>
    public class FeatureLocalCurvePatternData : SolidDnaObject<ILocalCurvePatternFeatureData>
    {
        public int PatternComponentCount => BaseObject.GetPatternComponentCount();
        public int SkippedItemCount => BaseObject.GetSkippedItemCount();
        public int D1TotalInstances => BaseObject.D1InstanceCount;
        public int D2TotalInstances => BaseObject.D2InstanceCount;
        public ModelFeature[] PatternComponentFeatures { protected set; get; }
        public bool IsDirection2Enabled => BaseObject.Dir2Specified;

        #region Constructor

        /// <summary>
        /// Default constructor
        /// </summary>
        public FeatureLocalCurvePatternData(object model) : base((ILocalCurvePatternFeatureData)model)
        {
            var llpSeedComps = (object[])BaseObject.PatternComponentArray;
            PatternComponentFeatures = new ModelFeature[llpSeedComps.Count()];

            for (int i = 0; i < llpSeedComps.Count(); i++)
            {
                PatternComponentFeatures[i] = new ModelFeature((Feature)llpSeedComps[i]);
            }
        }

        #endregion
    }
}
