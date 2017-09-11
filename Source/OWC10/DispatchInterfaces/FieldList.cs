using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.OWC10Api
{
	/// <summary>
	/// DispatchInterface FieldList 
	/// SupportByVersion OWC10, 1
	/// </summary>
	[SupportByVersion("OWC10", 1)]
	[EntityType(EntityType.IsDispatchInterface), BaseType]
 	public class FieldList : COMObject
	{
		#pragma warning disable

		#region Type Information

		/// <summary>
		/// Instance Type
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Advanced), Browsable(false), Category("NetOffice"), CoreOverridden]
		public override Type InstanceType
		{
			get
			{
				return LateBindingApiWrapperType;
			}
		}

        private static Type _type;

		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public static Type LateBindingApiWrapperType
        {
            get
            {
                if (null == _type)
                    _type = typeof(FieldList);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public FieldList(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public FieldList(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public FieldList(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public FieldList(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public FieldList(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public FieldList(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public FieldList() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public FieldList(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public Int32 ClipboardFormat
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "ClipboardFormat");
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public string InstanceID
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "InstanceID");
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1), NativeResult]
		public stdole.IFont Font
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Font", paramsArray);
                return returnItem as stdole.IFont;
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "Font", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public bool MultiSelect
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "MultiSelect");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "MultiSelect", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public NetOffice.OWC10Api.Enums.FieldListSelectRestriction SelectRestriction
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OWC10Api.Enums.FieldListSelectRestriction>(this, "SelectRestriction");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "SelectRestriction", value);
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="bVisible">bool bVisible</param>
		[SupportByVersion("OWC10", 1)]
		public NetOffice.OWC10Api.FieldListHierarchy CreateHierarchy(bool bVisible)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OWC10Api.FieldListHierarchy>(this, "CreateHierarchy", NetOffice.OWC10Api.FieldListHierarchy.LateBindingApiWrapperType, bVisible);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="iWidth">Int32 iWidth</param>
		/// <param name="iHeight">Int32 iHeight</param>
		/// <param name="pip">stdole.IPicture pip</param>
		/// <param name="crMask">Int32 crMask</param>
		[SupportByVersion("OWC10", 1)]
		public Int32 AddBitmap(Int32 iWidth, Int32 iHeight, stdole.IPicture pip, Int32 crMask)
		{
			return Factory.ExecuteInt32MethodGet(this, "AddBitmap", iWidth, iHeight, pip, crMask);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="pfln">NetOffice.OWC10Api.FieldListNode pfln</param>
		[SupportByVersion("OWC10", 1)]
		public NetOffice.OWC10Api.FieldListNode GetNextSelected(NetOffice.OWC10Api.FieldListNode pfln)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OWC10Api.FieldListNode>(this, "GetNextSelected", NetOffice.OWC10Api.FieldListNode.LateBindingApiWrapperType, pfln);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public void ClearSelection()
		{
			 Factory.ExecuteMethod(this, "ClearSelection");
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="iImage">Int32 iImage</param>
		/// <param name="iOverlay">Int32 iOverlay</param>
		[SupportByVersion("OWC10", 1)]
		public void SetOverlayImage(Int32 iImage, Int32 iOverlay)
		{
			 Factory.ExecuteMethod(this, "SetOverlayImage", iImage, iOverlay);
		}

		#endregion

		#pragma warning restore
	}
}
