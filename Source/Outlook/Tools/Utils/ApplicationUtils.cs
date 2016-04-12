using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NetOffice.OutlookApi.Tools.Utils
{
    /// <summary>
    /// Application related utils
    /// </summary>
    public class ApplicationUtils
    {
        #region Fields

        private CommonUtils _owner;

        #endregion

        #region Ctor

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="owner">owner instance</param>
        protected internal ApplicationUtils(CommonUtils owner)
        {
            if (null == owner)
                throw new ArgumentNullException("owner");
            _owner = owner;
        }

        #endregion

        #region Properties

        /// <summary>
        /// Host application is currently visible
        /// </summary>
        public bool Visible
        {
            get
            {
                OutlookDialogUtils utils = _owner.Dialog as OutlookDialogUtils;
                if (null != utils)
                    return utils.HostVisible;
                else
                    return false;
            }
        }

        #endregion
    }
}
