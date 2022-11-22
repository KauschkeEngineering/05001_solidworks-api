using SolidWorks.Interop.sldworks;

namespace AngelSix.SolidDna
{
    /// <summary>
    /// Represents a SolidWorks selection manager
    /// </summary>
    public class ConfigurationManager : SolidDnaObject<IConfigurationManager>
    {

        /// <summary>
        /// Contains the current active configuration information
        /// </summary>
        public ModelConfiguration ActiveConfiguration { get; protected set; } = null;

        #region Constructor

        /// <summary>
        /// Default constructor
        /// </summary>
        public ConfigurationManager(IConfigurationManager model) : base(model)
        {
            ActiveConfiguration = new ModelConfiguration(BaseObject.ActiveConfiguration);
        }

        #endregion

    }
}
