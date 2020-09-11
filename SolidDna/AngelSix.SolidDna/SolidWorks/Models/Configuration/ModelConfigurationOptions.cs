using SolidWorks.Interop.swconst;

namespace AngelSix.SolidDna
{
    /// <summary>
    /// Configuration options.
    /// <see cref="swInConfigurationOpts_e"/>
    /// Represents a configuration option used in multiple calls
    /// 
    /// NOTE: Known types are here http://help.solidworks.com/2020/english/api/swconst/SOLIDWORKS.Interop.swconst~SOLIDWORKS.Interop.swconst.swInConfigurationOpts_e.html
    /// 
    /// </summary>
    public enum ModelConfigurationOptions
    {
        /// <summary>
        /// Suppress Features in Configuration Property
        /// </summary>
        ConfigPropertySuppressFeatures = swInConfigurationOpts_e.swConfigPropertySuppressFeatures,

        /// <summary>
        /// This configuration
        /// </summary>
        ThisConfiguration = swInConfigurationOpts_e.swThisConfiguration,

        /// <summary>
        /// All configurations
        /// </summary>
        AllConfiguration = swInConfigurationOpts_e.swAllConfiguration,

        /// <summary>
        /// Specific configuration
        /// </summary>
        SpecificConfiguration = swInConfigurationOpts_e.swSpecifyConfiguration,

        /// <summary>
        /// Linked to parent
        /// </summary>
        /// <remarks>
        ///     Valid only for derived configurations;
        ///     if specified for non-derived configurations, 
        ///     then the active configuration is used
        /// </remarks>
        LinkedToParent = swInConfigurationOpts_e.swLinkedToParent,

        /// <summary>
        /// Speedpak Configuration
        /// </summary>
        SpeedpakConfiguration = swInConfigurationOpts_e.swSpeedpakConfiguration
    }
}
