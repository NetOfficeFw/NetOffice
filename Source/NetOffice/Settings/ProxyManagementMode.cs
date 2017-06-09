namespace NetOffice
{
    /// <summary>
    /// Represents the enumeration of all proxy management modes.
    /// </summary>
    public enum ProxyManagementMode
    {
        /// <summary>
        /// Represents the mode in which COM proxies are tracked using the object instance.
        /// </summary>
        Default,

        /// <summary>
        /// Represents the mode in which COM proxies are tracked using weak references.
        /// </summary>
        Weak
    }
}
