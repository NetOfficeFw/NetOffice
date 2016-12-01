using System;
using NetRuntimeSystem = System;
using System.Runtime.InteropServices;
using System.Runtime.CompilerServices;
using System.ComponentModel;
using System.Reflection;
using System.Collections.Generic;
using NetOffice;
namespace NetOffice.PowerPointApi
{
	///<summary>
	/// DispatchInterface PPDialogs 
	/// SupportByVersion PowerPoint, 9
	///</summary>
	[SupportByVersionAttribute("PowerPoint", 9)]
	[EntityTypeAttribute(EntityType.IsDispatchInterface)]
	public class PPDialogs : Collection
	{
		#pragma warning disable
		#region Type Information

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
        
		#region Construction

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
		
		/// <param name="progId">registered ProgID</param>
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
		[SupportByVersionAttribute("PowerPoint", 9)]
		public NetOffice.PowerPointApi.Application Application
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Application", paramsArray);
				NetOffice.PowerPointApi.Application newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.PowerPointApi.Application.LateBindingApiWrapperType) as NetOffice.PowerPointApi.Application;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9
		/// Get
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 9)]
		public NetOffice.PowerPointApi.Tags Tags
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Tags", paramsArray);
				NetOffice.PowerPointApi.Tags newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.PowerPointApi.Tags.LateBindingApiWrapperType) as NetOffice.PowerPointApi.Tags;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 9)]
		public string Name
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Name", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "Name", paramsArray);
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion PowerPoint 9
		/// 
		/// </summary>
		/// <param name="index">object index</param>
		[SupportByVersionAttribute("PowerPoint", 9)]
		[NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item")]
		public NetOffice.PowerPointApi.PPDialog this[object index]
		{
			get
			{
				object[] paramsArray = Invoker.ValidateParamsArray(index);
				object returnItem = Invoker.MethodReturn(this, "Item", paramsArray);
				NetOffice.PowerPointApi.PPDialog newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.PowerPointApi.PPDialog.LateBindingApiWrapperType) as NetOffice.PowerPointApi.PPDialog;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9
		/// 
		/// </summary>
		/// <param name="left">Single Left</param>
		/// <param name="top">Single Top</param>
		/// <param name="width">Single Width</param>
		/// <param name="height">Single Height</param>
		/// <param name="modal">optional NetOffice.OfficeApi.Enums.MsoTriState Modal = -1</param>
		/// <param name="parentWindow">optional object ParentWindow = null (Nothing in visual basic)</param>
		/// <param name="position">optional NetOffice.PowerPointApi.Enums.PpDialogPositioning Position = 1</param>
		/// <param name="displayHelp">optional NetOffice.OfficeApi.Enums.MsoTriState DisplayHelp = -1</param>
		[SupportByVersionAttribute("PowerPoint", 9)]
		public NetOffice.PowerPointApi.PPDialog AddDialog(Single left, Single top, Single width, Single height, object modal, object parentWindow, object position, object displayHelp)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(left, top, width, height, modal, parentWindow, position, displayHelp);
			object returnItem = Invoker.MethodReturn(this, "AddDialog", paramsArray);
			NetOffice.PowerPointApi.PPDialog newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.PowerPointApi.PPDialog.LateBindingApiWrapperType) as NetOffice.PowerPointApi.PPDialog;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9
		/// 
		/// </summary>
		/// <param name="left">Single Left</param>
		/// <param name="top">Single Top</param>
		/// <param name="width">Single Width</param>
		/// <param name="height">Single Height</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("PowerPoint", 9)]
		public NetOffice.PowerPointApi.PPDialog AddDialog(Single left, Single top, Single width, Single height)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(left, top, width, height);
			object returnItem = Invoker.MethodReturn(this, "AddDialog", paramsArray);
			NetOffice.PowerPointApi.PPDialog newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.PowerPointApi.PPDialog.LateBindingApiWrapperType) as NetOffice.PowerPointApi.PPDialog;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9
		/// 
		/// </summary>
		/// <param name="left">Single Left</param>
		/// <param name="top">Single Top</param>
		/// <param name="width">Single Width</param>
		/// <param name="height">Single Height</param>
		/// <param name="modal">optional NetOffice.OfficeApi.Enums.MsoTriState Modal = -1</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("PowerPoint", 9)]
		public NetOffice.PowerPointApi.PPDialog AddDialog(Single left, Single top, Single width, Single height, object modal)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(left, top, width, height, modal);
			object returnItem = Invoker.MethodReturn(this, "AddDialog", paramsArray);
			NetOffice.PowerPointApi.PPDialog newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.PowerPointApi.PPDialog.LateBindingApiWrapperType) as NetOffice.PowerPointApi.PPDialog;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9
		/// 
		/// </summary>
		/// <param name="left">Single Left</param>
		/// <param name="top">Single Top</param>
		/// <param name="width">Single Width</param>
		/// <param name="height">Single Height</param>
		/// <param name="modal">optional NetOffice.OfficeApi.Enums.MsoTriState Modal = -1</param>
		/// <param name="parentWindow">optional object ParentWindow = null (Nothing in visual basic)</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("PowerPoint", 9)]
		public NetOffice.PowerPointApi.PPDialog AddDialog(Single left, Single top, Single width, Single height, object modal, object parentWindow)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(left, top, width, height, modal, parentWindow);
			object returnItem = Invoker.MethodReturn(this, "AddDialog", paramsArray);
			NetOffice.PowerPointApi.PPDialog newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.PowerPointApi.PPDialog.LateBindingApiWrapperType) as NetOffice.PowerPointApi.PPDialog;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9
		/// 
		/// </summary>
		/// <param name="left">Single Left</param>
		/// <param name="top">Single Top</param>
		/// <param name="width">Single Width</param>
		/// <param name="height">Single Height</param>
		/// <param name="modal">optional NetOffice.OfficeApi.Enums.MsoTriState Modal = -1</param>
		/// <param name="parentWindow">optional object ParentWindow = null (Nothing in visual basic)</param>
		/// <param name="position">optional NetOffice.PowerPointApi.Enums.PpDialogPositioning Position = 1</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("PowerPoint", 9)]
		public NetOffice.PowerPointApi.PPDialog AddDialog(Single left, Single top, Single width, Single height, object modal, object parentWindow, object position)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(left, top, width, height, modal, parentWindow, position);
			object returnItem = Invoker.MethodReturn(this, "AddDialog", paramsArray);
			NetOffice.PowerPointApi.PPDialog newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.PowerPointApi.PPDialog.LateBindingApiWrapperType) as NetOffice.PowerPointApi.PPDialog;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9
		/// 
		/// </summary>
		/// <param name="left">Single Left</param>
		/// <param name="top">Single Top</param>
		/// <param name="width">Single Width</param>
		/// <param name="height">Single Height</param>
		/// <param name="modal">optional NetOffice.OfficeApi.Enums.MsoTriState Modal = -1</param>
		/// <param name="parentWindow">optional object ParentWindow = null (Nothing in visual basic)</param>
		/// <param name="position">optional NetOffice.PowerPointApi.Enums.PpDialogPositioning Position = 1</param>
		/// <param name="displayHelp">optional NetOffice.OfficeApi.Enums.MsoTriState DisplayHelp = -1</param>
		[SupportByVersionAttribute("PowerPoint", 9)]
		public NetOffice.PowerPointApi.PPDialog AddTabDialog(Single left, Single top, Single width, Single height, object modal, object parentWindow, object position, object displayHelp)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(left, top, width, height, modal, parentWindow, position, displayHelp);
			object returnItem = Invoker.MethodReturn(this, "AddTabDialog", paramsArray);
			NetOffice.PowerPointApi.PPDialog newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.PowerPointApi.PPDialog.LateBindingApiWrapperType) as NetOffice.PowerPointApi.PPDialog;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9
		/// 
		/// </summary>
		/// <param name="left">Single Left</param>
		/// <param name="top">Single Top</param>
		/// <param name="width">Single Width</param>
		/// <param name="height">Single Height</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("PowerPoint", 9)]
		public NetOffice.PowerPointApi.PPDialog AddTabDialog(Single left, Single top, Single width, Single height)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(left, top, width, height);
			object returnItem = Invoker.MethodReturn(this, "AddTabDialog", paramsArray);
			NetOffice.PowerPointApi.PPDialog newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.PowerPointApi.PPDialog.LateBindingApiWrapperType) as NetOffice.PowerPointApi.PPDialog;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9
		/// 
		/// </summary>
		/// <param name="left">Single Left</param>
		/// <param name="top">Single Top</param>
		/// <param name="width">Single Width</param>
		/// <param name="height">Single Height</param>
		/// <param name="modal">optional NetOffice.OfficeApi.Enums.MsoTriState Modal = -1</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("PowerPoint", 9)]
		public NetOffice.PowerPointApi.PPDialog AddTabDialog(Single left, Single top, Single width, Single height, object modal)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(left, top, width, height, modal);
			object returnItem = Invoker.MethodReturn(this, "AddTabDialog", paramsArray);
			NetOffice.PowerPointApi.PPDialog newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.PowerPointApi.PPDialog.LateBindingApiWrapperType) as NetOffice.PowerPointApi.PPDialog;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9
		/// 
		/// </summary>
		/// <param name="left">Single Left</param>
		/// <param name="top">Single Top</param>
		/// <param name="width">Single Width</param>
		/// <param name="height">Single Height</param>
		/// <param name="modal">optional NetOffice.OfficeApi.Enums.MsoTriState Modal = -1</param>
		/// <param name="parentWindow">optional object ParentWindow = null (Nothing in visual basic)</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("PowerPoint", 9)]
		public NetOffice.PowerPointApi.PPDialog AddTabDialog(Single left, Single top, Single width, Single height, object modal, object parentWindow)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(left, top, width, height, modal, parentWindow);
			object returnItem = Invoker.MethodReturn(this, "AddTabDialog", paramsArray);
			NetOffice.PowerPointApi.PPDialog newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.PowerPointApi.PPDialog.LateBindingApiWrapperType) as NetOffice.PowerPointApi.PPDialog;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9
		/// 
		/// </summary>
		/// <param name="left">Single Left</param>
		/// <param name="top">Single Top</param>
		/// <param name="width">Single Width</param>
		/// <param name="height">Single Height</param>
		/// <param name="modal">optional NetOffice.OfficeApi.Enums.MsoTriState Modal = -1</param>
		/// <param name="parentWindow">optional object ParentWindow = null (Nothing in visual basic)</param>
		/// <param name="position">optional NetOffice.PowerPointApi.Enums.PpDialogPositioning Position = 1</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("PowerPoint", 9)]
		public NetOffice.PowerPointApi.PPDialog AddTabDialog(Single left, Single top, Single width, Single height, object modal, object parentWindow, object position)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(left, top, width, height, modal, parentWindow, position);
			object returnItem = Invoker.MethodReturn(this, "AddTabDialog", paramsArray);
			NetOffice.PowerPointApi.PPDialog newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.PowerPointApi.PPDialog.LateBindingApiWrapperType) as NetOffice.PowerPointApi.PPDialog;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9
		/// 
		/// </summary>
		/// <param name="resourceDLL">string resourceDLL</param>
		/// <param name="nResID">Int32 nResID</param>
		/// <param name="bModal">optional NetOffice.OfficeApi.Enums.MsoTriState bModal = -1</param>
		/// <param name="parentWindow">optional object ParentWindow = null (Nothing in visual basic)</param>
		/// <param name="position">optional NetOffice.PowerPointApi.Enums.PpDialogPositioning Position = 1</param>
		[SupportByVersionAttribute("PowerPoint", 9)]
		public NetOffice.PowerPointApi.PPDialog LoadDialog(string resourceDLL, Int32 nResID, object bModal, object parentWindow, object position)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(resourceDLL, nResID, bModal, parentWindow, position);
			object returnItem = Invoker.MethodReturn(this, "LoadDialog", paramsArray);
			NetOffice.PowerPointApi.PPDialog newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.PowerPointApi.PPDialog.LateBindingApiWrapperType) as NetOffice.PowerPointApi.PPDialog;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9
		/// 
		/// </summary>
		/// <param name="resourceDLL">string resourceDLL</param>
		/// <param name="nResID">Int32 nResID</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("PowerPoint", 9)]
		public NetOffice.PowerPointApi.PPDialog LoadDialog(string resourceDLL, Int32 nResID)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(resourceDLL, nResID);
			object returnItem = Invoker.MethodReturn(this, "LoadDialog", paramsArray);
			NetOffice.PowerPointApi.PPDialog newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.PowerPointApi.PPDialog.LateBindingApiWrapperType) as NetOffice.PowerPointApi.PPDialog;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9
		/// 
		/// </summary>
		/// <param name="resourceDLL">string resourceDLL</param>
		/// <param name="nResID">Int32 nResID</param>
		/// <param name="bModal">optional NetOffice.OfficeApi.Enums.MsoTriState bModal = -1</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("PowerPoint", 9)]
		public NetOffice.PowerPointApi.PPDialog LoadDialog(string resourceDLL, Int32 nResID, object bModal)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(resourceDLL, nResID, bModal);
			object returnItem = Invoker.MethodReturn(this, "LoadDialog", paramsArray);
			NetOffice.PowerPointApi.PPDialog newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.PowerPointApi.PPDialog.LateBindingApiWrapperType) as NetOffice.PowerPointApi.PPDialog;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9
		/// 
		/// </summary>
		/// <param name="resourceDLL">string resourceDLL</param>
		/// <param name="nResID">Int32 nResID</param>
		/// <param name="bModal">optional NetOffice.OfficeApi.Enums.MsoTriState bModal = -1</param>
		/// <param name="parentWindow">optional object ParentWindow = null (Nothing in visual basic)</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("PowerPoint", 9)]
		public NetOffice.PowerPointApi.PPDialog LoadDialog(string resourceDLL, Int32 nResID, object bModal, object parentWindow)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(resourceDLL, nResID, bModal, parentWindow);
			object returnItem = Invoker.MethodReturn(this, "LoadDialog", paramsArray);
			NetOffice.PowerPointApi.PPDialog newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.PowerPointApi.PPDialog.LateBindingApiWrapperType) as NetOffice.PowerPointApi.PPDialog;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9
		/// 
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 9)]
		public NetOffice.PowerPointApi.PPAlert AddAlert()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "AddAlert", paramsArray);
			NetOffice.PowerPointApi.PPAlert newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.PowerPointApi.PPAlert.LateBindingApiWrapperType) as NetOffice.PowerPointApi.PPAlert;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9
		/// 
		/// </summary>
		/// <param name="text">string Text</param>
		/// <param name="type">NetOffice.PowerPointApi.Enums.PpAlertType Type</param>
		/// <param name="icon">NetOffice.PowerPointApi.Enums.PpAlertIcon icon</param>
		/// <param name="parentWindow">optional object ParentWindow = null (Nothing in visual basic)</param>
		[SupportByVersionAttribute("PowerPoint", 9)]
		public NetOffice.PowerPointApi.Enums.PpAlertButton RunCharacterAlert(string text, NetOffice.PowerPointApi.Enums.PpAlertType type, NetOffice.PowerPointApi.Enums.PpAlertIcon icon, object parentWindow)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(text, type, icon, parentWindow);
			object returnItem = Invoker.MethodReturn(this, "RunCharacterAlert", paramsArray);
			int intReturnItem = NetRuntimeSystem.Convert.ToInt32(returnItem);
			return (NetOffice.PowerPointApi.Enums.PpAlertButton)intReturnItem;
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9
		/// 
		/// </summary>
		/// <param name="text">string Text</param>
		/// <param name="type">NetOffice.PowerPointApi.Enums.PpAlertType Type</param>
		/// <param name="icon">NetOffice.PowerPointApi.Enums.PpAlertIcon icon</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("PowerPoint", 9)]
		public NetOffice.PowerPointApi.Enums.PpAlertButton RunCharacterAlert(string text, NetOffice.PowerPointApi.Enums.PpAlertType type, NetOffice.PowerPointApi.Enums.PpAlertIcon icon)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(text, type, icon);
			object returnItem = Invoker.MethodReturn(this, "RunCharacterAlert", paramsArray);
			int intReturnItem = NetRuntimeSystem.Convert.ToInt32(returnItem);
			return (NetOffice.PowerPointApi.Enums.PpAlertButton)intReturnItem;
		}

		#endregion
		#pragma warning restore
	}
}