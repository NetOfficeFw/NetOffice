using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using NetOffice;
using NetOffice.Attributes;

namespace NetOffice.OfficeApi.Events
{	
	#pragma warning disable
	
	#region SinkPoint Interface

	[SupportByVersion("Office", 12,14,15,16)]
    [InternalEntity(InternalEntityKind.ComEventInterface)]
    [ComImport, Guid("000CDB07-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch), TypeLibType((short)0x1010)]
	public interface _CustomXMLPartEvents
	{
		[SupportByVersion("Office", 12,14,15,16)]
        [SinkArgument("newNode", typeof(OfficeApi.CustomXMLNode))]
        [SinkArgument("inUndoRedo", SinkArgumentType.Bool)]       
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1)]
		void NodeAfterInsert([In, MarshalAs(UnmanagedType.IDispatch)] object newNode, [In] object inUndoRedo);

		[SupportByVersion("Office", 12,14,15,16)]
        [SinkArgument("oldNode", typeof(OfficeApi.CustomXMLNode))]
        [SinkArgument("oldParentNode", typeof(OfficeApi.CustomXMLNode))]
        [SinkArgument("oldNextSibling", typeof(OfficeApi.CustomXMLNode))]
        [SinkArgument("inUndoRedo", SinkArgumentType.Bool)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2)]
		void NodeAfterDelete([In, MarshalAs(UnmanagedType.IDispatch)] object oldNode, [In, MarshalAs(UnmanagedType.IDispatch)] object oldParentNode, [In, MarshalAs(UnmanagedType.IDispatch)] object oldNextSibling, [In] object inUndoRedo);

		[SupportByVersion("Office", 12,14,15,16)]

        [SinkArgument("oldNode", typeof(OfficeApi.CustomXMLNode))]
        [SinkArgument("newNode", typeof(OfficeApi.CustomXMLNode))]
        [SinkArgument("inUndoRedo", SinkArgumentType.Bool)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(3)]
		void NodeAfterReplace([In, MarshalAs(UnmanagedType.IDispatch)] object oldNode, [In, MarshalAs(UnmanagedType.IDispatch)] object newNode, [In] object inUndoRedo);
	}

    #endregion

    #region SinkHelper

    [InternalEntity(InternalEntityKind.SinkHelper)]
    [ComVisible(true), ClassInterface(ClassInterfaceType.None), TypeLibType(TypeLibTypeFlags.FHidden)]
	public class _CustomXMLPartEvents_SinkHelper : SinkHelper, _CustomXMLPartEvents
	{
		#region Static
		
		public static readonly string Id = "000CDB07-0000-0000-C000-000000000046";
		
		#endregion
	
		#region Construction

		public _CustomXMLPartEvents_SinkHelper(ICOMObject eventClass, IConnectionPoint connectPoint): base(eventClass)
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

			NetOffice.OfficeApi.CustomXMLNode newNewNode = Factory.CreateKnownObjectFromComProxy<NetOffice.OfficeApi.CustomXMLNode>(EventClass, newNode, NetOffice.OfficeApi.CustomXMLNode.LateBindingApiWrapperType);
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

            NetOffice.OfficeApi.CustomXMLNode newOldNode = Factory.CreateKnownObjectFromComProxy<NetOffice.OfficeApi.CustomXMLNode>(EventClass, oldNode, NetOffice.OfficeApi.CustomXMLNode.LateBindingApiWrapperType);
            NetOffice.OfficeApi.CustomXMLNode newOldParentNode = Factory.CreateKnownObjectFromComProxy<NetOffice.OfficeApi.CustomXMLNode>(EventClass, oldParentNode, NetOffice.OfficeApi.CustomXMLNode.LateBindingApiWrapperType);
            NetOffice.OfficeApi.CustomXMLNode newOldNextSibling = Factory.CreateKnownObjectFromComProxy<NetOffice.OfficeApi.CustomXMLNode>(EventClass, oldNextSibling, NetOffice.OfficeApi.CustomXMLNode.LateBindingApiWrapperType);
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

            NetOffice.OfficeApi.CustomXMLNode newOldNode = Factory.CreateKnownObjectFromComProxy<NetOffice.OfficeApi.CustomXMLNode>(EventClass, oldNode, NetOffice.OfficeApi.CustomXMLNode.LateBindingApiWrapperType);
            NetOffice.OfficeApi.CustomXMLNode newNewNode = Factory.CreateKnownObjectFromComProxy<NetOffice.OfficeApi.CustomXMLNode>(EventClass, newNode, NetOffice.OfficeApi.CustomXMLNode.LateBindingApiWrapperType);
            bool newInUndoRedo = ToBoolean(inUndoRedo);
			object[] paramsArray = new object[3];
			paramsArray[0] = newOldNode;
			paramsArray[1] = newNewNode;
			paramsArray[2] = newInUndoRedo;
			EventBinding.RaiseCustomEvent("NodeAfterReplace", ref paramsArray);
		}

		#endregion
	}
	
	#endregion
	
	#pragma warning restore
}