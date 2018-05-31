using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using NetOffice;
using NetOffice.Attributes;

namespace NetOffice.OfficeApi.Behind.EventContracts
{
    [InternalEntity(InternalEntityKind.SinkHelper)]
    [ComVisible(true), ClassInterface(ClassInterfaceType.None), TypeLibType(TypeLibTypeFlags.FHidden)]
    public class _CustomXMLPartEvents_SinkHelper : SinkHelper, NetOffice.OfficeApi.EventContracts._CustomXMLPartEvents
    {
        #region Static

        public static readonly string Id = "000CDB07-0000-0000-C000-000000000046";

        #endregion

        #region Construction

        public _CustomXMLPartEvents_SinkHelper(ICOMObject eventClass, IConnectionPoint connectPoint) : base(eventClass)
        {
            SetupEventBinding(connectPoint);
        }

        #endregion

        #region _CustomXMLPartEvents

        public void NodeAfterInsert([In, MarshalAs(UnmanagedType.IDispatch)] object newNode, [In] object inUndoRedo)
        {
            if (!Validate("NodeAfterInsert"))
            {
                Invoker.ReleaseParamsArray(newNode, inUndoRedo);
                return;
            }

            NetOffice.OfficeApi.CustomXMLNode newNewNode = Factory.CreateKnownObjectFromComProxy<NetOffice.OfficeApi.CustomXMLNode>(EventClass, newNode, typeof(NetOffice.OfficeApi.CustomXMLNode));
            bool newInUndoRedo = ToBoolean(inUndoRedo);
            object[] paramsArray = new object[2];
            paramsArray[0] = newNewNode;
            paramsArray[1] = newInUndoRedo;
            EventBinding.RaiseCustomEvent("NodeAfterInsert", ref paramsArray);
        }

        public void NodeAfterDelete([In, MarshalAs(UnmanagedType.IDispatch)] object oldNode, [In, MarshalAs(UnmanagedType.IDispatch)] object oldParentNode, [In, MarshalAs(UnmanagedType.IDispatch)] object oldNextSibling, [In] object inUndoRedo)
        {
            if (!Validate("NodeAfterDelete"))
            {
                Invoker.ReleaseParamsArray(oldNode, oldParentNode, oldNextSibling, inUndoRedo);
                return;
            }

            NetOffice.OfficeApi.CustomXMLNode newOldNode = Factory.CreateKnownObjectFromComProxy<NetOffice.OfficeApi.CustomXMLNode>(EventClass, oldNode, typeof(NetOffice.OfficeApi.CustomXMLNode));
            NetOffice.OfficeApi.CustomXMLNode newOldParentNode = Factory.CreateKnownObjectFromComProxy<NetOffice.OfficeApi.CustomXMLNode>(EventClass, oldParentNode, typeof(NetOffice.OfficeApi.CustomXMLNode));
            NetOffice.OfficeApi.CustomXMLNode newOldNextSibling = Factory.CreateKnownObjectFromComProxy<NetOffice.OfficeApi.CustomXMLNode>(EventClass, oldNextSibling, typeof(NetOffice.OfficeApi.CustomXMLNode));
            bool newInUndoRedo = ToBoolean(inUndoRedo);
            object[] paramsArray = new object[4];
            paramsArray[0] = newOldNode;
            paramsArray[1] = newOldParentNode;
            paramsArray[2] = newOldNextSibling;
            paramsArray[3] = newInUndoRedo;
            EventBinding.RaiseCustomEvent("NodeAfterDelete", ref paramsArray);
        }

        public void NodeAfterReplace([In, MarshalAs(UnmanagedType.IDispatch)] object oldNode, [In, MarshalAs(UnmanagedType.IDispatch)] object newNode, [In] object inUndoRedo)
        {
            if (!Validate("NodeAfterReplace"))
            {
                Invoker.ReleaseParamsArray(oldNode, newNode, inUndoRedo); return;
            }

            NetOffice.OfficeApi.CustomXMLNode newOldNode = Factory.CreateKnownObjectFromComProxy<NetOffice.OfficeApi.CustomXMLNode>(EventClass, oldNode, typeof(NetOffice.OfficeApi.CustomXMLNode));
            NetOffice.OfficeApi.CustomXMLNode newNewNode = Factory.CreateKnownObjectFromComProxy<NetOffice.OfficeApi.CustomXMLNode>(EventClass, newNode, typeof(NetOffice.OfficeApi.CustomXMLNode));
            bool newInUndoRedo = ToBoolean(inUndoRedo);
            object[] paramsArray = new object[3];
            paramsArray[0] = newOldNode;
            paramsArray[1] = newNewNode;
            paramsArray[2] = newInUndoRedo;
            EventBinding.RaiseCustomEvent("NodeAfterReplace", ref paramsArray);
        }

        #endregion
    }
}
