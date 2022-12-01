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

    }
}
