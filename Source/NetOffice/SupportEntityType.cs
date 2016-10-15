namespace NetOffice
{
    /// <summary>
    /// Search parameter for the EntityIsAvailable Method
    /// </summary>
    public enum SupportEntityType
    {
        /// <summary>
        /// looking for a method or a property
        /// </summary>
        Both = 0,

        /// <summary>
        /// looking for a property only
        /// </summary>
        Property = 1,

        /// <summary>
        /// looking for a method only
        /// </summary>
        Method = 2
    }
}