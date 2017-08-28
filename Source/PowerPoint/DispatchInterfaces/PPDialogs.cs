using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.PowerPointApi
{
	/// <summary>
	/// DispatchInterface PPDialogs 
	/// SupportByVersion PowerPoint, 9
	/// </summary>
	[SupportByVersion("PowerPoint", 9)]
	[EntityType(EntityType.IsDispatchInterface), HasIndexProperty(IndexInvoke.Method, "Item")]
 	public class PPDialogs : Collection
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
                    _type = typeof(PPDialogs);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public PPDialogs(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public PPDialogs(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public PPDialogs(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public PPDialogs(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public PPDialogs(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public PPDialogs(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public PPDialogs() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public PPDialogs(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion PowerPoint 9
		/// Get
		/// </summary>
		[SupportByVersion("PowerPoint", 9)]
		public NetOffice.PowerPointApi.Application Application
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.PowerPointApi.Application>(this, "Application", NetOffice.PowerPointApi.Application.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9
		/// Get
		/// </summary>
		[SupportByVersion("PowerPoint", 9)]
		public NetOffice.PowerPointApi.Tags Tags
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.PowerPointApi.Tags>(this, "Tags", NetOffice.PowerPointApi.Tags.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9
		/// Get/Set
		/// </summary>
		[SupportByVersion("PowerPoint", 9)]
		public string Name
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Name");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Name", value);
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion PowerPoint 9
		/// </summary>
		/// <param name="index">object index</param>
		[SupportByVersion("PowerPoint", 9)]
		[NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item"), IndexProperty]
		public NetOffice.PowerPointApi.PPDialog this[object index]
		{
			get
			{
				return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.PPDialog>(this, "Item", NetOffice.PowerPointApi.PPDialog.LateBindingApiWrapperType, index);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9
		/// </summary>
		/// <param name="left">Single left</param>
		/// <param name="top">Single top</param>
		/// <param name="width">Single width</param>
		/// <param name="height">Single height</param>
		/// <param name="modal">optional NetOffice.OfficeApi.Enums.MsoTriState Modal = -1</param>
		/// <param name="parentWindow">optional object ParentWindow = null (Nothing in visual basic)</param>
		/// <param name="position">optional NetOffice.PowerPointApi.Enums.PpDialogPositioning Position = 1</param>
		/// <param name="displayHelp">optional NetOffice.OfficeApi.Enums.MsoTriState DisplayHelp = -1</param>
		[SupportByVersion("PowerPoint", 9)]
		public NetOffice.PowerPointApi.PPDialog AddDialog(Single left, Single top, Single width, Single height, object modal, object parentWindow, object position, object displayHelp)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.PPDialog>(this, "AddDialog", NetOffice.PowerPointApi.PPDialog.LateBindingApiWrapperType, new object[]{ left, top, width, height, modal, parentWindow, position, displayHelp });
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9
		/// </summary>
		/// <param name="left">Single left</param>
		/// <param name="top">Single top</param>
		/// <param name="width">Single width</param>
		/// <param name="height">Single height</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 9)]
		public NetOffice.PowerPointApi.PPDialog AddDialog(Single left, Single top, Single width, Single height)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.PPDialog>(this, "AddDialog", NetOffice.PowerPointApi.PPDialog.LateBindingApiWrapperType, left, top, width, height);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9
		/// </summary>
		/// <param name="left">Single left</param>
		/// <param name="top">Single top</param>
		/// <param name="width">Single width</param>
		/// <param name="height">Single height</param>
		/// <param name="modal">optional NetOffice.OfficeApi.Enums.MsoTriState Modal = -1</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 9)]
		public NetOffice.PowerPointApi.PPDialog AddDialog(Single left, Single top, Single width, Single height, object modal)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.PPDialog>(this, "AddDialog", NetOffice.PowerPointApi.PPDialog.LateBindingApiWrapperType, new object[]{ left, top, width, height, modal });
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9
		/// </summary>
		/// <param name="left">Single left</param>
		/// <param name="top">Single top</param>
		/// <param name="width">Single width</param>
		/// <param name="height">Single height</param>
		/// <param name="modal">optional NetOffice.OfficeApi.Enums.MsoTriState Modal = -1</param>
		/// <param name="parentWindow">optional object ParentWindow = null (Nothing in visual basic)</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 9)]
		public NetOffice.PowerPointApi.PPDialog AddDialog(Single left, Single top, Single width, Single height, object modal, object parentWindow)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.PPDialog>(this, "AddDialog", NetOffice.PowerPointApi.PPDialog.LateBindingApiWrapperType, new object[]{ left, top, width, height, modal, parentWindow });
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9
		/// </summary>
		/// <param name="left">Single left</param>
		/// <param name="top">Single top</param>
		/// <param name="width">Single width</param>
		/// <param name="height">Single height</param>
		/// <param name="modal">optional NetOffice.OfficeApi.Enums.MsoTriState Modal = -1</param>
		/// <param name="parentWindow">optional object ParentWindow = null (Nothing in visual basic)</param>
		/// <param name="position">optional NetOffice.PowerPointApi.Enums.PpDialogPositioning Position = 1</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 9)]
		public NetOffice.PowerPointApi.PPDialog AddDialog(Single left, Single top, Single width, Single height, object modal, object parentWindow, object position)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.PPDialog>(this, "AddDialog", NetOffice.PowerPointApi.PPDialog.LateBindingApiWrapperType, new object[]{ left, top, width, height, modal, parentWindow, position });
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9
		/// </summary>
		/// <param name="left">Single left</param>
		/// <param name="top">Single top</param>
		/// <param name="width">Single width</param>
		/// <param name="height">Single height</param>
		/// <param name="modal">optional NetOffice.OfficeApi.Enums.MsoTriState Modal = -1</param>
		/// <param name="parentWindow">optional object ParentWindow = null (Nothing in visual basic)</param>
		/// <param name="position">optional NetOffice.PowerPointApi.Enums.PpDialogPositioning Position = 1</param>
		/// <param name="displayHelp">optional NetOffice.OfficeApi.Enums.MsoTriState DisplayHelp = -1</param>
		[SupportByVersion("PowerPoint", 9)]
		public NetOffice.PowerPointApi.PPDialog AddTabDialog(Single left, Single top, Single width, Single height, object modal, object parentWindow, object position, object displayHelp)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.PPDialog>(this, "AddTabDialog", NetOffice.PowerPointApi.PPDialog.LateBindingApiWrapperType, new object[]{ left, top, width, height, modal, parentWindow, position, displayHelp });
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9
		/// </summary>
		/// <param name="left">Single left</param>
		/// <param name="top">Single top</param>
		/// <param name="width">Single width</param>
		/// <param name="height">Single height</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 9)]
		public NetOffice.PowerPointApi.PPDialog AddTabDialog(Single left, Single top, Single width, Single height)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.PPDialog>(this, "AddTabDialog", NetOffice.PowerPointApi.PPDialog.LateBindingApiWrapperType, left, top, width, height);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9
		/// </summary>
		/// <param name="left">Single left</param>
		/// <param name="top">Single top</param>
		/// <param name="width">Single width</param>
		/// <param name="height">Single height</param>
		/// <param name="modal">optional NetOffice.OfficeApi.Enums.MsoTriState Modal = -1</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 9)]
		public NetOffice.PowerPointApi.PPDialog AddTabDialog(Single left, Single top, Single width, Single height, object modal)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.PPDialog>(this, "AddTabDialog", NetOffice.PowerPointApi.PPDialog.LateBindingApiWrapperType, new object[]{ left, top, width, height, modal });
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9
		/// </summary>
		/// <param name="left">Single left</param>
		/// <param name="top">Single top</param>
		/// <param name="width">Single width</param>
		/// <param name="height">Single height</param>
		/// <param name="modal">optional NetOffice.OfficeApi.Enums.MsoTriState Modal = -1</param>
		/// <param name="parentWindow">optional object ParentWindow = null (Nothing in visual basic)</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 9)]
		public NetOffice.PowerPointApi.PPDialog AddTabDialog(Single left, Single top, Single width, Single height, object modal, object parentWindow)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.PPDialog>(this, "AddTabDialog", NetOffice.PowerPointApi.PPDialog.LateBindingApiWrapperType, new object[]{ left, top, width, height, modal, parentWindow });
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9
		/// </summary>
		/// <param name="left">Single left</param>
		/// <param name="top">Single top</param>
		/// <param name="width">Single width</param>
		/// <param name="height">Single height</param>
		/// <param name="modal">optional NetOffice.OfficeApi.Enums.MsoTriState Modal = -1</param>
		/// <param name="parentWindow">optional object ParentWindow = null (Nothing in visual basic)</param>
		/// <param name="position">optional NetOffice.PowerPointApi.Enums.PpDialogPositioning Position = 1</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 9)]
		public NetOffice.PowerPointApi.PPDialog AddTabDialog(Single left, Single top, Single width, Single height, object modal, object parentWindow, object position)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.PPDialog>(this, "AddTabDialog", NetOffice.PowerPointApi.PPDialog.LateBindingApiWrapperType, new object[]{ left, top, width, height, modal, parentWindow, position });
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9
		/// </summary>
		/// <param name="resourceDLL">string resourceDLL</param>
		/// <param name="nResID">Int32 nResID</param>
		/// <param name="bModal">optional NetOffice.OfficeApi.Enums.MsoTriState bModal = -1</param>
		/// <param name="parentWindow">optional object ParentWindow = null (Nothing in visual basic)</param>
		/// <param name="position">optional NetOffice.PowerPointApi.Enums.PpDialogPositioning Position = 1</param>
		[SupportByVersion("PowerPoint", 9)]
		public NetOffice.PowerPointApi.PPDialog LoadDialog(string resourceDLL, Int32 nResID, object bModal, object parentWindow, object position)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.PPDialog>(this, "LoadDialog", NetOffice.PowerPointApi.PPDialog.LateBindingApiWrapperType, new object[]{ resourceDLL, nResID, bModal, parentWindow, position });
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9
		/// </summary>
		/// <param name="resourceDLL">string resourceDLL</param>
		/// <param name="nResID">Int32 nResID</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 9)]
		public NetOffice.PowerPointApi.PPDialog LoadDialog(string resourceDLL, Int32 nResID)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.PPDialog>(this, "LoadDialog", NetOffice.PowerPointApi.PPDialog.LateBindingApiWrapperType, resourceDLL, nResID);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9
		/// </summary>
		/// <param name="resourceDLL">string resourceDLL</param>
		/// <param name="nResID">Int32 nResID</param>
		/// <param name="bModal">optional NetOffice.OfficeApi.Enums.MsoTriState bModal = -1</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 9)]
		public NetOffice.PowerPointApi.PPDialog LoadDialog(string resourceDLL, Int32 nResID, object bModal)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.PPDialog>(this, "LoadDialog", NetOffice.PowerPointApi.PPDialog.LateBindingApiWrapperType, resourceDLL, nResID, bModal);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9
		/// </summary>
		/// <param name="resourceDLL">string resourceDLL</param>
		/// <param name="nResID">Int32 nResID</param>
		/// <param name="bModal">optional NetOffice.OfficeApi.Enums.MsoTriState bModal = -1</param>
		/// <param name="parentWindow">optional object ParentWindow = null (Nothing in visual basic)</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 9)]
		public NetOffice.PowerPointApi.PPDialog LoadDialog(string resourceDLL, Int32 nResID, object bModal, object parentWindow)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.PPDialog>(this, "LoadDialog", NetOffice.PowerPointApi.PPDialog.LateBindingApiWrapperType, resourceDLL, nResID, bModal, parentWindow);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9
		/// </summary>
		[SupportByVersion("PowerPoint", 9)]
		public NetOffice.PowerPointApi.PPAlert AddAlert()
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.PPAlert>(this, "AddAlert", NetOffice.PowerPointApi.PPAlert.LateBindingApiWrapperType);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9
		/// </summary>
		/// <param name="text">string text</param>
		/// <param name="type">NetOffice.PowerPointApi.Enums.PpAlertType type</param>
		/// <param name="icon">NetOffice.PowerPointApi.Enums.PpAlertIcon icon</param>
		/// <param name="parentWindow">optional object ParentWindow = null (Nothing in visual basic)</param>
		[SupportByVersion("PowerPoint", 9)]
		public NetOffice.PowerPointApi.Enums.PpAlertButton RunCharacterAlert(string text, NetOffice.PowerPointApi.Enums.PpAlertType type, NetOffice.PowerPointApi.Enums.PpAlertIcon icon, object parentWindow)
		{
			return Factory.ExecuteEnumMethodGet<NetOffice.PowerPointApi.Enums.PpAlertButton>(this, "RunCharacterAlert", text, type, icon, parentWindow);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9
		/// </summary>
		/// <param name="text">string text</param>
		/// <param name="type">NetOffice.PowerPointApi.Enums.PpAlertType type</param>
		/// <param name="icon">NetOffice.PowerPointApi.Enums.PpAlertIcon icon</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 9)]
		public NetOffice.PowerPointApi.Enums.PpAlertButton RunCharacterAlert(string text, NetOffice.PowerPointApi.Enums.PpAlertType type, NetOffice.PowerPointApi.Enums.PpAlertIcon icon)
		{
			return Factory.ExecuteEnumMethodGet<NetOffice.PowerPointApi.Enums.PpAlertButton>(this, "RunCharacterAlert", text, type, icon);
		}

		#endregion

		#pragma warning restore
	}
}
