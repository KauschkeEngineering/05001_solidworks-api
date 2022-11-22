using SolidWorks.Interop.sldworks;

namespace AngelSix.SolidDna
{
    /// <summary>
    /// Represents a SolidWorks model configuration
    /// </summary>
    public class ModelConfiguration : SolidDnaObject<Configuration>
    {

        #region Properties

        public string Name => BaseObject.Name;

        public bool IsDirty => BaseObject.IsDirty();

        #endregion

        #region Constructor

        /// <summary>
        /// Default constructor
        /// </summary>
        public ModelConfiguration(Configuration model) : base(model)
        {

        }

        #endregion

        public Component GetRootComponent(bool resolve)
        {
            return BaseObject != null ? new Component(BaseObject.GetRootComponent3(resolve)) : null;
        }
    }
}
