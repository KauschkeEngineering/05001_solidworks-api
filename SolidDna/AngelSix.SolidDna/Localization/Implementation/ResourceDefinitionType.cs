namespace AngelSix.SolidDna
{
    /// <summary>
    /// The type of resource used in a localization manager
    /// </summary>
    public enum ResourceDefinitionType
    {
        /// <summary>
        /// The resource is embedded in the assembly
        /// </summary>
        EmbeddedResource = 1,

        /// <summary>
        /// The resource is on the file system of the application
        /// </summary>
        File = 2,

        /// <summary>
        /// The resource is a remote URL
        /// </summary>
        Url = 3
    }
}
