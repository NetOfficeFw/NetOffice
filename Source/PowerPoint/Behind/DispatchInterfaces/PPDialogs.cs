using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.PowerPointApi;

namespace NetOffice.PowerPointApi.Behind
{
	/// <summary>
	/// DispatchInterface PPDialogs 
	/// SupportByVersion PowerPoint, 9
	/// </summary>
	[SupportByVersion("PowerPoint", 9)]
	[EntityType(EntityType.IsDispatchInterface), HasIndexProperty(IndexInvoke.Method, "Item")]
 	public class PPDialogs : Collection, NetOffice.PowerPointApi.PPDialogs
	{
		#pragma warning disable

		#region Type Information

        /// <summary>
        /// Contract Type
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Advanced), Browsable(false), Category("NetOffice"), CoreOverridden]
        public override Type ContractType
        {
            get
            {
                if(null == _contractType)
                    _contractType = typeof(NetOffice.PowerPointApi.PPDialogs);
                return _contractType;
            }
        }
        private static Type _contractType;


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

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public PPDialogs() : base()
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
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.PowerPointApi.Application>(this, "Application", typeof(NetOffice.PowerPointApi.Application));
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
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.PowerPointApi.Tags>(this, "Tags", typeof(NetOffice.PowerPointApi.Tags));
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
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "Name");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Name", value);
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
				return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.PPDialog>(this, "Item", typeof(NetOffice.PowerPointApi.PPDialog), index);
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
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.PPDialog>(this, "AddDialog", typeof(NetOffice.PowerPointApi.PPDialog), new object[]{ left, top, width, height, modal, parentWindow, position, displayHelp });
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
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.PPDialog>(this, "AddDialog", typeof(NetOffice.PowerPointApi.PPDialog), left, top, width, height);
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
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.PPDialog>(this, "AddDialog", typeof(NetOffice.PowerPointApi.PPDialog), new object[]{ left, top, width, height, modal });
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
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.PPDialog>(this, "AddDialog", typeof(NetOffice.PowerPointApi.PPDialog), new object[]{ left, top, width, height, modal, parentWindow });
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
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.PPDialog>(this, "AddDialog", typeof(NetOffice.PowerPointApi.PPDialog), new object[]{ left, top, width, height, modal, parentWindow, position });
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
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.PPDialog>(this, "AddTabDialog", typeof(NetOffice.PowerPointApi.PPDialog), new object[]{ left, top, width, height, modal, parentWindow, position, displayHelp });
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
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.PPDialog>(this, "AddTabDialog", typeof(NetOffice.PowerPointApi.PPDialog), left, top, width, height);
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
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.PPDialog>(this, "AddTabDialog", typeof(NetOffice.PowerPointApi.PPDialog), new object[]{ left, top, width, height, modal });
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
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.PPDialog>(this, "AddTabDialog", typeof(NetOffice.PowerPointApi.PPDialog), new object[]{ left, top, width, height, modal, parentWindow });
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
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.PPDialog>(this, "AddTabDialog", typeof(NetOffice.PowerPointApi.PPDialog), new object[]{ left, top, width, height, modal, parentWindow, position });
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
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.PPDialog>(this, "LoadDialog", typeof(NetOffice.PowerPointApi.PPDialog), new object[]{ resourceDLL, nResID, bModal, parentWindow, position });
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
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.PPDialog>(this, "LoadDialog", typeof(NetOffice.PowerPointApi.PPDialog), resourceDLL, nResID);
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
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.PPDialog>(this, "LoadDialog", typeof(NetOffice.PowerPointApi.PPDialog), resourceDLL, nResID, bModal);
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
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.PPDialog>(this, "LoadDialog", typeof(NetOffice.PowerPointApi.PPDialog), resourceDLL, nResID, bModal, parentWindow);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9
		/// </summary>
		[SupportByVersion("PowerPoint", 9)]
		public NetOffice.PowerPointApi.PPAlert AddAlert()
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.PPAlert>(this, "AddAlert", typeof(NetOffice.PowerPointApi.PPAlert));
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
			return InvokerService.InvokeInternal.ExecuteEnumMethodGet<NetOffice.PowerPointApi.Enums.PpAlertButton>(this, "RunCharacterAlert", text, type, icon, parentWindow);
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
			return InvokerService.InvokeInternal.ExecuteEnumMethodGet<NetOffice.PowerPointApi.Enums.PpAlertButton>(this, "RunCharacterAlert", text, type, icon);
		}

		#endregion

		#pragma warning restore
	}
}


