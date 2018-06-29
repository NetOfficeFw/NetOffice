using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NetOffice.ComTypes
{
    /// <summary>
    /// Contains the most common interface id's related to COM and Microsoft Office
    /// </summary>
    public class WellKnownIID
    {
        /// <summary>
        /// 00000000-0000-0000-C000-000000000046
        /// </summary>
        public Guid IID_IUnkown = new Guid("00000000-0000-0000-C000-000000000046");

        /// <summary>
        /// 00020400-0000-0000-c000-000000000046
        /// </summary>
        public Guid IID_IDispatch = new Guid("00020400-0000-0000-c000-000000000046");

        /// <summary>
        /// B65AD801-ABAF-11D0-BB8B-00A0C90F2744
        /// </summary>
        public Guid IID_Extensibility2 = new Guid("B65AD801-ABAF-11D0-BB8B-00A0C90F2744");

        /// <summary>
        /// 000C0396-0000-0000-C000-000000000046
        /// </summary>
        public Guid IID_IRibbonExtensibility = new Guid("000C0396-0000-0000-C000-000000000046");

        /// <summary>
        /// 000C033E-0000-0000-C000-000000000046
        /// </summary>
        public Guid IID_ICustomTaskPaneConsumer = new Guid("000C033E-0000-0000-C000-000000000046");

        /// <summary>
        /// 000CD706-0000-0000-C000-000000000046
        /// </summary>
        public Guid IID_IDocumentInspector= new Guid("000CD706-0000-0000-C000-000000000046");

        /// <summary>
        /// 000C03C4-0000-0000-C000-000000000046
        /// </summary>
        public Guid IID_IBlogExtensibility = new Guid("000C03C4-0000-0000-C000-000000000046");

        /// <summary>
        /// 000C03C5-0000-0000-C000-000000000046
        /// </summary>
        public Guid IID_IBlogPictureExtensibility = new Guid("000C03C5-0000-0000-C000-000000000046");

        /// <summary>
        /// 000CD6A3-0000-0000-C000-000000000046
        /// </summary>
        public Guid IID_SignatureProvider = new Guid("000CD6A3-0000-0000-C000-000000000046");

        /// <summary>
        /// 00063059-0000-0000-C000-000000000046"
        /// </summary>
        public Guid IID_FormRegionStartup = new Guid("00063059-0000-0000-C000-000000000046");

        /// <summary>
        /// 0006307E-0000-0000-C000-000000000046
        /// </summary>
        public Guid IID_PropertyPage = new Guid("0006307E-0000-0000-C000-000000000046");

        /// <summary>
        /// 0006307F-0000-0000-C000-000000000046
        /// </summary>
        public Guid IID_PropertyPageSite = new Guid("0006307F-0000-0000-C000-000000000046");

        /// <summary>
        /// EC0E6191-DB51-11D3-8F3E-00C04F3651B8
        /// </summary>
        public Guid IID_IRtdServer = new Guid("EC0E6191-DB51-11D3-8F3E-00C04F3651B8");

        /// <summary>
        /// Returns the IID constant for an id if its well known
        /// </summary>
        /// <param name="id">target id</param>
        /// <param name="faulty">result if not found or null/empty to return id argument as string</param>
        /// <returns>iid or id string</returns>
        public string GetIID(Guid id, string faulty = null)
        {
            string result = String.IsNullOrWhiteSpace(faulty) ? id.ToString() : faulty;

            if (IID_IUnkown == id)
                return "IID_IUnkown";
            else if (IID_IDispatch == id)
                return "IID_IDispatch";
            else if (IID_Extensibility2 == id)
                return "IID_Extensibility2";
            else if (IID_IRibbonExtensibility == id)
                return "IID_IRibbonExtensibility";
            else if (IID_ICustomTaskPaneConsumer == id)
                return "IID_ICustomTaskPaneConsumer";
            else if (IID_IDocumentInspector == id)
                return "IID_IDocumentInspector";
            else if (IID_IDispatch == id)
                return "IID_IBlogExtensibility";
            else if (IID_IBlogPictureExtensibility == id)
                return "IID_IBlogPictureExtensibility";
            else if (IID_SignatureProvider == id)
                return "IID_SignatureProvider";
            else if (IID_IDispatch == id)
                return "IID_FormRegionStartup";
            else if (IID_PropertyPage == id)
                return "IID_PropertyPage";
            else if (IID_PropertyPageSite == id)
                return "IID_PropertyPageSite";
            else if (IID_IDispatch == id)
                return "IID_IRtdServer";

            return result;
        }
    }
}
