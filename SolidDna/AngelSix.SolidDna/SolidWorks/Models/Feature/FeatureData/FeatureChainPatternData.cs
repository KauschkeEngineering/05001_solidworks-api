using SolidWorks.Interop.sldworks;

namespace AngelSix.SolidDna
{
    /// <summary>
    /// Represents a SolidWorks Chain Pattern feature data
    /// </summary>
    public class FeatureChainPatternData : SolidDnaObject<IChainPatternFeatureData>
    {
        public int InstanceCount => BaseObject.InstanceCount;
        public ModelFeature Group1PatternFeature { protected set; get; }
        public ModelFeature Group2PatternFeature { protected set; get; }

        #region Constructor

        /// <summary>
        /// Default constructor
        /// </summary>
        public FeatureChainPatternData(object model) : base((IChainPatternFeatureData)model)
        {
            
        }

        #endregion

        public bool AccessSelections(Model TopLevelModel, Component componentToModify = null)
        {
            var status = BaseObject.AccessSelections(TopLevelModel.AsModelDoc2, componentToModify);
            Group1PatternFeature = new ModelFeature((Feature)BaseObject.Group1PatternComponent);
            Group2PatternFeature = new ModelFeature((Feature)BaseObject.Group2PatternComponent);
            return status;
        }

        public void ReleaseSelectionAccess()
        {
            BaseObject.ReleaseSelectionAccess();
            Group1PatternFeature.Dispose();
            Group1PatternFeature = null;
            Group2PatternFeature.Dispose();
            Group2PatternFeature = null;
        }

    }
}
