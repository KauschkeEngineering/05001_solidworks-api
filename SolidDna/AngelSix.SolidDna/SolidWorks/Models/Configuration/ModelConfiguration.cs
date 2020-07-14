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

        #endregion

        #region Constructor

        /// <summary>
        /// Default constructor
        /// </summary>
        public ModelConfiguration(Configuration model) : base(model)
        {

        }

        #endregion
    }
}
