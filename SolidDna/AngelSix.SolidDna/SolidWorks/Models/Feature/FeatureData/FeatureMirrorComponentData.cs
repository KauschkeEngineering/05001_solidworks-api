using SolidWorks.Interop.sldworks;
using System.Linq;

namespace AngelSix.SolidDna
{
    /// <summary>
    /// Represents a SolidWorks Mirror Component feature data
    /// </summary>
    public class FeatureMirrorComponentData : SolidDnaObject<IMirrorComponentFeatureData>
    {
        public ModelFeature[] OppositeHandComponentFeatures { protected set; get; }

        #region Constructor

        /// <summary>
        /// Default constructor
        /// </summary>
        public FeatureMirrorComponentData(object model) : base((IMirrorComponentFeatureData)model)
        {
            var llpSeedComps = (object[])BaseObject.OppositeHandComponents;
            OppositeHandComponentFeatures = new ModelFeature[llpSeedComps.Count()];

            for (int i = 0; i < llpSeedComps.Count(); i++)
            {
                OppositeHandComponentFeatures[i] = new ModelFeature((Feature)llpSeedComps[i]);
            }
        }

        #endregion
    }
}