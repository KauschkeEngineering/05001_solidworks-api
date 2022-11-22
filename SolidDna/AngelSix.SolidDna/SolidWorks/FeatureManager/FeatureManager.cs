using SolidWorks.Interop.sldworks;

namespace AngelSix.SolidDna
{
    /// <summary>
    /// Represents a SolidWorks feature manager
    /// </summary>
    public class FeatureManager : SolidDnaObject<IFeatureManager>
    {
        public bool EnableFeatureTree
        {
            get => BaseObject.EnableFeatureTree;
            set => BaseObject.EnableFeatureTree = value;
        }
        public bool EnableFeatureTreeWindow
        {
            get => BaseObject.EnableFeatureTreeWindow;
            set => BaseObject.EnableFeatureTreeWindow = value;
        }

        #region Constructor

        /// <summary>
        /// Default constructor
        /// </summary>
        public FeatureManager(IFeatureManager model) : base(model)
        {

        }

        #endregion

        private ModelFeature CreateSpecificFeature(object featureData)
        {
            return new ModelFeature(BaseObject.CreateFeature(featureData));
        }

        public ModelFeature CreateFeature(object featureData)
        {
            return CreateSpecificFeature(featureData);
        }

        public ModelFeature CreateFeature(FeatureChainPatternData featureData)
        {
            return CreateSpecificFeature(featureData);
        }

        public ModelFeature CreateFeature(FeatureLinearPatternData featureData)
        {
            return CreateSpecificFeature(featureData);
        }

    }
}
