using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NetOffice.OutlookApi.Tools.Contribution
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

        #region Methods

        /// <summary>
        /// Try get PropertyPageSite(container) from PropertyPage
        /// </summary>
        /// <param name="page">target page</param>
        /// <returns>page container</returns>
        public static NetOffice.OutlookApi.Native.PropertyPageSite TryGetPageContainer(Native.PropertyPage page)
        {
            try
            {
                Type myType = typeof(object);
                string assembly = System.Text.RegularExpressions.Regex.Replace(myType.Assembly.CodeBase, "mscorlib.dll", "System.Windows.Forms.dll");
                assembly = System.Text.RegularExpressions.Regex.Replace(assembly, "file:///", "");
                assembly = System.Reflection.AssemblyName.GetAssemblyName(assembly).FullName;
                Type unmanaged = Type.GetType(System.Reflection.Assembly.CreateQualifiedName(assembly, "System.Windows.Forms.UnsafeNativeMethods"));
                Type oleObj = unmanaged.GetNestedType("IOleObject");
                System.Reflection.MethodInfo mi = oleObj.GetMethod("GetClientSite");
                object myppSite = mi.Invoke(page, null);
                return myppSite as Native.PropertyPageSite;
            }
            catch
            {
                return null;
            }           
        }

        #endregion
    }
}