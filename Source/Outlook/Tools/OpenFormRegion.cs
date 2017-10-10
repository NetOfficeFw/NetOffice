using System;

namespace NetOffice.OutlookApi.Tools
{
    /// <summary>
    /// Represents an open FormRegion
    /// </summary>
    public class OpenFormRegion
    {
        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="underlyingRegion">origin formregion</param>
        public OpenFormRegion(FormRegion underlyingRegion)
        {
            UnderlyingRegion = underlyingRegion;
            UnderlyingRegion.CloseEvent += UnderlyingRegion_CloseEvent;
        }

        internal event Action<OpenFormRegion> Close;

        /// <summary>
        /// Origin FormRegion
        /// </summary>
        public FormRegion UnderlyingRegion { get; private set; }

        /// <summary>
        /// Tag as any
        /// </summary>
        public object Tag { get; set; }

        private void UnderlyingRegion_CloseEvent()
        {
            Close?.Invoke(this);
            UnderlyingRegion.CloseEvent -= UnderlyingRegion_CloseEvent;
        }
    }
}
