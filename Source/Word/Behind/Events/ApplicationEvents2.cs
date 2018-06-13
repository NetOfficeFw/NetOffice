using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using NetOffice;
using NetOffice.Attributes;

namespace NetOffice.WordApi.Behind.EventContracts
{
    /// <summary>
    /// Default implementation of <see cref="NetOffice.WordApi.EventContracts.ApplicationEvents2"/>
    /// </summary>
    [InternalEntity(InternalEntityKind.SinkHelper)]
    [ComVisible(true), ClassInterface(ClassInterfaceType.None), TypeLibType(TypeLibTypeFlags.FHidden)]
    public class ApplicationEvents2_SinkHelper : SinkHelper, NetOffice.WordApi.EventContracts.ApplicationEvents2
    {
        #region Static

        /// <summary>
        /// Interface Id from ApplicationEvents2
        /// </summary>
        public static readonly string Id = "000209FE-0000-0000-C000-000000000046";

        #endregion

        #region Ctor

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="eventClass"></param>
        /// <param name="connectPoint"></param>
        public ApplicationEvents2_SinkHelper(ICOMObject eventClass, IConnectionPoint connectPoint) : base(eventClass)
        {
            SetupEventBinding(connectPoint);
        }

        #endregion

        #region ApplicationEvents2

        /// <summary>
        /// 
        /// </summary>
        public void Startup()
        {
            if (!Validate("Startup"))
            {
                return;
            }

            object[] paramsArray = new object[0];
            EventBinding.RaiseCustomEvent("Startup", ref paramsArray);
        }

        /// <summary>
        /// 
        /// </summary>
        public void Quit()
        {
            if (!Validate("Quit"))
            {
                return;
            }

            object[] paramsArray = new object[0];
            EventBinding.RaiseCustomEvent("Quit", ref paramsArray);        }

        /// <summary>
        /// 
        /// </summary>
        public void DocumentChange()
        {
            if (!Validate("DocumentChange"))
            {
                return;
            }

            object[] paramsArray = new object[0];
            EventBinding.RaiseCustomEvent("DocumentChange", ref paramsArray);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="doc"></param>
        public void DocumentOpen([In, MarshalAs(UnmanagedType.IDispatch)] object doc)
        {
            if (!Validate("DocumentOpen"))
            {
                return;
            }

            NetOffice.WordApi.Document newDoc = Factory.CreateKnownObjectFromComProxy<NetOffice.WordApi.Document>(EventClass, doc, typeof(NetOffice.WordApi.Document));
            object[] paramsArray = new object[1];
            paramsArray[0] = newDoc;
            EventBinding.RaiseCustomEvent("DocumentOpen", ref paramsArray);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="doc"></param>
        /// <param name="cancel"></param>
        public void DocumentBeforeClose([In, MarshalAs(UnmanagedType.IDispatch)] object doc, [In] [Out] ref object cancel)
        {
            if (!Validate("DocumentBeforeClose"))
            {
                Invoker.ReleaseParamsArray(doc, cancel);
                return;
            }

            NetOffice.WordApi.Document newDoc = Factory.CreateKnownObjectFromComProxy<NetOffice.WordApi.Document>(EventClass, doc, typeof(NetOffice.WordApi.Document));
            object[] paramsArray = new object[2];
            paramsArray[0] = newDoc;
            paramsArray.SetValue(cancel, 1);
            EventBinding.RaiseCustomEvent("DocumentBeforeClose", ref paramsArray);

            cancel = ToBoolean(paramsArray[1]);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="doc"></param>
        /// <param name="cancel"></param>
        public void DocumentBeforePrint([In, MarshalAs(UnmanagedType.IDispatch)] object doc, [In] [Out] ref object cancel)
        {
            if (!Validate("DocumentBeforePrint"))
            {
                Invoker.ReleaseParamsArray(doc, cancel);
                return;
            }

            NetOffice.WordApi.Document newDoc = Factory.CreateKnownObjectFromComProxy<NetOffice.WordApi.Document>(EventClass, doc, typeof(NetOffice.WordApi.Document));
            object[] paramsArray = new object[2];
            paramsArray[0] = newDoc;
            paramsArray.SetValue(cancel, 1);
            EventBinding.RaiseCustomEvent("DocumentBeforePrint", ref paramsArray);

            cancel = ToBoolean(paramsArray[1]);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="doc"></param>
        /// <param name="saveAsUI"></param>
        /// <param name="cancel"></param>
        public void DocumentBeforeSave([In, MarshalAs(UnmanagedType.IDispatch)] object doc, [In] [Out] ref object saveAsUI, [In] [Out] ref object cancel)
        {
            if (!Validate("DocumentBeforeSave"))
            {
                Invoker.ReleaseParamsArray(doc, saveAsUI, cancel);
                return;
            }

            NetOffice.WordApi.Document newDoc = Factory.CreateKnownObjectFromComProxy<NetOffice.WordApi.Document>(EventClass, doc, typeof(NetOffice.WordApi.Document));
            object[] paramsArray = new object[3];
            paramsArray[0] = newDoc;
            paramsArray.SetValue(saveAsUI, 1);
            paramsArray.SetValue(cancel, 2);
            EventBinding.RaiseCustomEvent("DocumentBeforeSave", ref paramsArray);

            saveAsUI = ToBoolean(paramsArray[1]);
            cancel = ToBoolean(paramsArray[2]);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="doc"></param>
        public void NewDocument([In, MarshalAs(UnmanagedType.IDispatch)] object doc)
        {
            if (!Validate("NewDocument"))
            {
                Invoker.ReleaseParamsArray(doc);
                return;
            }

            NetOffice.WordApi.Document newDoc = Factory.CreateKnownObjectFromComProxy<NetOffice.WordApi.Document>(EventClass, doc, typeof(NetOffice.WordApi.Document));
            object[] paramsArray = new object[1];
            paramsArray[0] = newDoc;
            EventBinding.RaiseCustomEvent("NewDocument", ref paramsArray);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="doc"></param>
        /// <param name="wn"></param>
        public void WindowActivate([In, MarshalAs(UnmanagedType.IDispatch)] object doc, [In, MarshalAs(UnmanagedType.IDispatch)] object wn)
        {
            if (!Validate("WindowActivate"))
            {
                Invoker.ReleaseParamsArray(doc, wn);
                return;
            }

            NetOffice.WordApi.Document newDoc = Factory.CreateKnownObjectFromComProxy<NetOffice.WordApi.Document>(EventClass, doc, typeof(NetOffice.WordApi.Document));
            NetOffice.WordApi.Window newWn = Factory.CreateKnownObjectFromComProxy<NetOffice.WordApi.Window>(EventClass, wn, typeof(NetOffice.WordApi.Window));
            object[] paramsArray = new object[2];
            paramsArray[0] = newDoc;
            paramsArray[1] = newWn;
            EventBinding.RaiseCustomEvent("WindowActivate", ref paramsArray);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="doc"></param>
        /// <param name="wn"></param>
        public void WindowDeactivate([In, MarshalAs(UnmanagedType.IDispatch)] object doc, [In, MarshalAs(UnmanagedType.IDispatch)] object wn)
        {
            if (!Validate("WindowDeactivate"))
            {
                Invoker.ReleaseParamsArray(doc, wn);
                return;
            }

            NetOffice.WordApi.Document newDoc = Factory.CreateKnownObjectFromComProxy<NetOffice.WordApi.Document>(EventClass, doc, typeof(NetOffice.WordApi.Document));
            NetOffice.WordApi.Window newWn = Factory.CreateKnownObjectFromComProxy<NetOffice.WordApi.Window>(EventClass, wn, typeof(NetOffice.WordApi.Window));
            object[] paramsArray = new object[2];
            paramsArray[0] = newDoc;
            paramsArray[1] = newWn;
            EventBinding.RaiseCustomEvent("WindowDeactivate", ref paramsArray);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sel"></param>
        public void WindowSelectionChange([In, MarshalAs(UnmanagedType.IDispatch)] object sel)
        {
            if (!Validate("WindowSelectionChange"))
            {
                Invoker.ReleaseParamsArray(sel);
                return;
            }

            NetOffice.WordApi.Selection newSel = Factory.CreateKnownObjectFromComProxy<NetOffice.WordApi.Selection>(EventClass, sel, typeof(NetOffice.WordApi.Selection));
            object[] paramsArray = new object[1];
            paramsArray[0] = newSel;
            EventBinding.RaiseCustomEvent("WindowSelectionChange", ref paramsArray);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sel"></param>
        /// <param name="cancel"></param>
        public void WindowBeforeRightClick([In, MarshalAs(UnmanagedType.IDispatch)] object sel, [In] [Out] ref object cancel)
        {
            if (!Validate("WindowBeforeRightClick"))
            {
                Invoker.ReleaseParamsArray(sel, cancel);
                return;
            }

            NetOffice.WordApi.Selection newSel = Factory.CreateKnownObjectFromComProxy<NetOffice.WordApi.Selection>(EventClass, sel, typeof(NetOffice.WordApi.Selection));
            object[] paramsArray = new object[2];
            paramsArray[0] = newSel;
            paramsArray.SetValue(cancel, 1);
            EventBinding.RaiseCustomEvent("WindowBeforeRightClick", ref paramsArray);

            cancel = ToBoolean(paramsArray[1]);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sel"></param>
        /// <param name="cancel"></param>
        public void WindowBeforeDoubleClick([In, MarshalAs(UnmanagedType.IDispatch)] object sel, [In] [Out] ref object cancel)
        {
            if (!Validate("WindowBeforeDoubleClick"))
            {
                Invoker.ReleaseParamsArray(sel, cancel);
                return;
            }

            NetOffice.WordApi.Selection newSel = Factory.CreateKnownObjectFromComProxy<NetOffice.WordApi.Selection>(EventClass, sel, typeof(NetOffice.WordApi.Selection));
            object[] paramsArray = new object[2];
            paramsArray[0] = newSel;
            paramsArray.SetValue(cancel, 1);
            EventBinding.RaiseCustomEvent("WindowBeforeDoubleClick", ref paramsArray);

            cancel = ToBoolean(paramsArray[1]);
        }

        #endregion
    }
}

