using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using NetOffice;
using NetOffice.Attributes;

namespace NetOffice.WordApi.Behind.EventContracts
{
    /// <summary>
    /// Default implementation of <see cref="NetOffice.WordApi.EventContracts.DocumentEvents2"/>
    /// </summary>
    [InternalEntity(InternalEntityKind.SinkHelper)]
    [ComVisible(true), ClassInterface(ClassInterfaceType.None), TypeLibType(TypeLibTypeFlags.FHidden)]
    public class DocumentEvents2_SinkHelper : SinkHelper, NetOffice.WordApi.EventContracts.DocumentEvents2
    {
        #region Static

        /// <summary>
        /// Interface Id from DocumentEvents2
        /// </summary>
        public static readonly string Id = "00020A02-0000-0000-C000-000000000046";

        #endregion

        #region Ctor

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="eventClass"></param>
        /// <param name="connectPoint"></param>
        public DocumentEvents2_SinkHelper(ICOMObject eventClass, IConnectionPoint connectPoint) : base(eventClass)
        {
            SetupEventBinding(connectPoint);
        }

        #endregion

        #region DocumentEvents2

        /// <summary>
        /// 
        /// </summary>
        public void New()
        {
            if (!Validate("New"))
            {
                return;
            }

            object[] paramsArray = new object[0];
            EventBinding.RaiseCustomEvent("New", ref paramsArray);
        }

        /// <summary>
        /// 
        /// </summary>
        public void Open()
        {
            if (!Validate("Open"))
            {
                return;
            }

            object[] paramsArray = new object[0];
            EventBinding.RaiseCustomEvent("Open", ref paramsArray);        }

        /// <summary>
        /// 
        /// </summary>
        public void Close()
        {
            if (!Validate("Close"))
            {
                return;
            }

            object[] paramsArray = new object[0];
            EventBinding.RaiseCustomEvent("Close", ref paramsArray);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="syncEventType"></param>
        public void Sync([In] object syncEventType)
        {
            if (!Validate("Sync"))
            {
                Invoker.ReleaseParamsArray(syncEventType);
                return;
            }

            NetOffice.OfficeApi.Enums.MsoSyncEventType newSyncEventType = (NetOffice.OfficeApi.Enums.MsoSyncEventType)syncEventType;
            object[] paramsArray = new object[1];
            paramsArray[0] = newSyncEventType;
            EventBinding.RaiseCustomEvent("Sync", ref paramsArray);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="newXMLNode"></param>
        /// <param name="inUndoRedo"></param>
        public void XMLAfterInsert([In, MarshalAs(UnmanagedType.IDispatch)] object newXMLNode, [In] object inUndoRedo)
        {
            if (!Validate("XMLAfterInsert"))
            {
                Invoker.ReleaseParamsArray(newXMLNode, inUndoRedo);
                return;
            }

            NetOffice.WordApi.XMLNode newNewXMLNode = Factory.CreateKnownObjectFromComProxy<NetOffice.WordApi.XMLNode>(EventClass, newXMLNode, typeof(NetOffice.WordApi.XMLNode));
            bool newInUndoRedo = ToBoolean(inUndoRedo);
            object[] paramsArray = new object[2];
            paramsArray[0] = newNewXMLNode;
            paramsArray[1] = newInUndoRedo;
            EventBinding.RaiseCustomEvent("XMLAfterInsert", ref paramsArray);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="deletedRange"></param>
        /// <param name="oldXMLNode"></param>
        /// <param name="inUndoRedo"></param>
        public void XMLBeforeDelete([In, MarshalAs(UnmanagedType.IDispatch)] object deletedRange, [In, MarshalAs(UnmanagedType.IDispatch)] object oldXMLNode, [In] object inUndoRedo)
        {
            if (!Validate("XMLBeforeDelete"))
            {
                Invoker.ReleaseParamsArray(deletedRange, oldXMLNode, inUndoRedo);
                return;
            }

            NetOffice.WordApi.Range newDeletedRange = Factory.CreateKnownObjectFromComProxy<NetOffice.WordApi.Range>(EventClass, deletedRange, typeof(NetOffice.WordApi.Range));
            NetOffice.WordApi.XMLNode newOldXMLNode = Factory.CreateKnownObjectFromComProxy<NetOffice.WordApi.XMLNode>(EventClass, oldXMLNode, typeof(NetOffice.WordApi.XMLNode));
            bool newInUndoRedo = Convert.ToBoolean(inUndoRedo);
            object[] paramsArray = new object[3];
            paramsArray[0] = newDeletedRange;
            paramsArray[1] = newOldXMLNode;
            paramsArray[2] = newInUndoRedo;
            EventBinding.RaiseCustomEvent("XMLBeforeDelete", ref paramsArray);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="newContentControl"></param>
        /// <param name="inUndoRedo"></param>
        public void ContentControlAfterAdd([In, MarshalAs(UnmanagedType.IDispatch)] object newContentControl, [In] object inUndoRedo)
        {
            if (!Validate("ContentControlAfterAdd"))
            {
                Invoker.ReleaseParamsArray(newContentControl, inUndoRedo);
                return;
            }

            NetOffice.WordApi.ContentControl newNewContentControl = Factory.CreateKnownObjectFromComProxy<NetOffice.WordApi.ContentControl>(EventClass, newContentControl, typeof(NetOffice.WordApi.ContentControl));
            bool newInUndoRedo = ToBoolean(inUndoRedo);
            object[] paramsArray = new object[2];
            paramsArray[0] = newNewContentControl;
            paramsArray[1] = newInUndoRedo;
            EventBinding.RaiseCustomEvent("ContentControlAfterAdd", ref paramsArray);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="oldContentControl"></param>
        /// <param name="inUndoRedo"></param>
        public void ContentControlBeforeDelete([In, MarshalAs(UnmanagedType.IDispatch)] object oldContentControl, [In] object inUndoRedo)
        {
            if (!Validate("ContentControlBeforeDelete"))
            {
                Invoker.ReleaseParamsArray(oldContentControl, inUndoRedo);
                return;
            }

            NetOffice.WordApi.ContentControl newOldContentControl = Factory.CreateKnownObjectFromComProxy<NetOffice.WordApi.ContentControl>(EventClass, oldContentControl, typeof(NetOffice.WordApi.ContentControl));
            bool newInUndoRedo = Convert.ToBoolean(inUndoRedo);
            object[] paramsArray = new object[2];
            paramsArray[0] = newOldContentControl;
            paramsArray[1] = newInUndoRedo;
            EventBinding.RaiseCustomEvent("ContentControlBeforeDelete", ref paramsArray);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="contentControl"></param>
        /// <param name="cancel"></param>
        public void ContentControlOnExit([In, MarshalAs(UnmanagedType.IDispatch)] object contentControl, [In] [Out] ref object cancel)
        {
            if (!Validate("ContentControlOnExit"))
            {
                Invoker.ReleaseParamsArray(contentControl, cancel);
                return;
            }

            NetOffice.WordApi.ContentControl newContentControl = Factory.CreateKnownObjectFromComProxy<NetOffice.WordApi.ContentControl>(EventClass, contentControl, typeof(NetOffice.WordApi.ContentControl));
            object[] paramsArray = new object[2];
            paramsArray[0] = newContentControl;
            paramsArray.SetValue(cancel, 1);
            EventBinding.RaiseCustomEvent("ContentControlOnExit", ref paramsArray);

            cancel = ToBoolean(paramsArray[1]);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="contentControl"></param>
        public void ContentControlOnEnter([In, MarshalAs(UnmanagedType.IDispatch)] object contentControl)
        {
            if (!Validate("ContentControlOnEnter"))
            {
                Invoker.ReleaseParamsArray(contentControl);
                return;
            }

            NetOffice.WordApi.ContentControl newContentControl = Factory.CreateKnownObjectFromComProxy<NetOffice.WordApi.ContentControl>(EventClass, contentControl, typeof(NetOffice.WordApi.ContentControl));
            object[] paramsArray = new object[1];
            paramsArray[0] = newContentControl;
            EventBinding.RaiseCustomEvent("ContentControlOnEnter", ref paramsArray);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="contentControl"></param>
        /// <param name="content"></param>
        public void ContentControlBeforeStoreUpdate([In, MarshalAs(UnmanagedType.IDispatch)] object contentControl, [In] [Out] ref object content)
        {
            if (!Validate("ContentControlBeforeStoreUpdate"))
            {
                Invoker.ReleaseParamsArray(contentControl, content);
                return;
            }

            NetOffice.WordApi.ContentControl newContentControl = Factory.CreateKnownObjectFromComProxy<NetOffice.WordApi.ContentControl>(EventClass, contentControl, typeof(NetOffice.WordApi.ContentControl));
            object[] paramsArray = new object[2];
            paramsArray[0] = newContentControl;
            paramsArray.SetValue(content, 1);
            EventBinding.RaiseCustomEvent("ContentControlBeforeStoreUpdate", ref paramsArray);

            content = ToString(paramsArray[1]);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="contentControl"></param>
        /// <param name="content"></param>
        public void ContentControlBeforeContentUpdate([In, MarshalAs(UnmanagedType.IDispatch)] object contentControl, [In] [Out] ref object content)
        {
            if (!Validate("ContentControlBeforeStoreUpdate"))
            {
                Invoker.ReleaseParamsArray(contentControl, content);
                return;
            }

            NetOffice.WordApi.ContentControl newContentControl = Factory.CreateKnownObjectFromComProxy<NetOffice.WordApi.ContentControl>(EventClass, contentControl, typeof(NetOffice.WordApi.ContentControl));
            object[] paramsArray = new object[2];
            paramsArray[0] = newContentControl;
            paramsArray.SetValue(content, 1);
            EventBinding.RaiseCustomEvent("ContentControlBeforeContentUpdate", ref paramsArray);

            content = ToString(paramsArray[1]);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="range"></param>
        /// <param name="name"></param>
        /// <param name="category"></param>
        /// <param name="blockType"></param>
        /// <param name="template"></param>
        public void BuildingBlockInsert([In, MarshalAs(UnmanagedType.IDispatch)] object range, [In] object name, [In] object category, [In] object blockType, [In] object template)
        {
            if (!Validate("BuildingBlockInsert"))
            {
                Invoker.ReleaseParamsArray(range, name, category, blockType, template);
                return;
            }

            NetOffice.WordApi.Range newRange = Factory.CreateKnownObjectFromComProxy<NetOffice.WordApi.Range>(EventClass, range, typeof(NetOffice.WordApi.Range));
            string newName = ToString(name);
            string newCategory = ToString(category);
            string newBlockType = ToString(blockType);
            string newTemplate = ToString(template);
            object[] paramsArray = new object[5];
            paramsArray[0] = newRange;
            paramsArray[1] = newName;
            paramsArray[2] = newCategory;
            paramsArray[3] = newBlockType;
            paramsArray[4] = newTemplate;
            EventBinding.RaiseCustomEvent("BuildingBlockInsert", ref paramsArray);
        }

        #endregion
    }
}

