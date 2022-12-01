using SolidWorks.Interop.sldworks;
using System.Linq;

namespace AngelSix.SolidDna
{
    /// <summary>
    /// Represents a SolidWorks Local Circular Pattern feature data
    /// </summary>
    public class FeatureLocalCircularPatternData : SolidDnaObject<ILocalCircularPatternFeatureData>
    {
        public int SeedComponentCount => BaseObject.GetSeedComponentCount();
        public int SkippedItemCount => BaseObject.GetSkippedItemCount();
        public int D1TotalInstances => BaseObject.TotalInstances;
        public int D2TotalInstances => BaseObject.TotalInstances2;
        public ModelFeature[] SeedComponentFeatures { protected set; get; }
        public bool IsDirection2Enabled => BaseObject.Direction2;

        #region Constructor

        /// <summary>
        /// Default constructor
        /// </summary>
        public FeatureLocalCircularPatternData(object model) : base((ILocalCircularPatternFeatureData)model)
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
