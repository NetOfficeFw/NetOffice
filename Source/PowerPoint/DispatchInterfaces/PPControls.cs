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
	/// DispatchInterface PPControls 
	/// SupportByVersion PowerPoint, 9
	///</summary>
	[SupportByVersionAttribute("PowerPoint", 9)]
	[EntityTypeAttribute(EntityType.IsDispatchInterface)]
	public class PPControls : Collection
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
                    _type = typeof(PPControls);
                    
                return _type;
            }
        }
        
        #endregion
        
		#region Construction

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public PPControls(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public PPControls(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public PPControls(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public PPControls(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public PPControls(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public PPControls() : base()
		{
		}
		
		/// <param name="progId">registered ProgID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public PPControls(string progId) : base(progId)
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
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("PowerPoint", 9)]
		public NetOffice.OfficeApi.Enums.MsoTriState Visible
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Visible", paramsArray);
				int intReturnItem = NetRuntimeSystem.Convert.ToInt32(returnItem);
				return (NetOffice.OfficeApi.Enums.MsoTriState)intReturnItem;
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "Visible", paramsArray);
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
		public NetOffice.PowerPointApi.PPControl this[object index]
		{
			get
			{
				object[] paramsArray = Invoker.ValidateParamsArray(index);
				object returnItem = Invoker.MethodReturn(this, "Item", paramsArray);
				NetOffice.PowerPointApi.PPControl newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.PowerPointApi.PPControl;
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
		[SupportByVersionAttribute("PowerPoint", 9)]
		public NetOffice.PowerPointApi.PPPushButton AddPushButton(Single left, Single top, Single width, Single height)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(left, top, width, height);
			object returnItem = Invoker.MethodReturn(this, "AddPushButton", paramsArray);
			NetOffice.PowerPointApi.PPPushButton newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.PowerPointApi.PPPushButton.LateBindingApiWrapperType) as NetOffice.PowerPointApi.PPPushButton;
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
		[SupportByVersionAttribute("PowerPoint", 9)]
		public NetOffice.PowerPointApi.PPToggleButton AddToggleButton(Single left, Single top, Single width, Single height)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(left, top, width, height);
			object returnItem = Invoker.MethodReturn(this, "AddToggleButton", paramsArray);
			NetOffice.PowerPointApi.PPToggleButton newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.PowerPointApi.PPToggleButton.LateBindingApiWrapperType) as NetOffice.PowerPointApi.PPToggleButton;
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
		[SupportByVersionAttribute("PowerPoint", 9)]
		public NetOffice.PowerPointApi.PPBitmapButton AddBitmapButton(Single left, Single top, Single width, Single height)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(left, top, width, height);
			object returnItem = Invoker.MethodReturn(this, "AddBitmapButton", paramsArray);
			NetOffice.PowerPointApi.PPBitmapButton newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.PowerPointApi.PPBitmapButton.LateBindingApiWrapperType) as NetOffice.PowerPointApi.PPBitmapButton;
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
		[SupportByVersionAttribute("PowerPoint", 9)]
		public NetOffice.PowerPointApi.PPListBox AddListBox(Single left, Single top, Single width, Single height)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(left, top, width, height);
			object returnItem = Invoker.MethodReturn(this, "AddListBox", paramsArray);
			NetOffice.PowerPointApi.PPListBox newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.PowerPointApi.PPListBox.LateBindingApiWrapperType) as NetOffice.PowerPointApi.PPListBox;
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
		[SupportByVersionAttribute("PowerPoint", 9)]
		public NetOffice.PowerPointApi.PPCheckBox AddCheckBox(Single left, Single top, Single width, Single height)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(left, top, width, height);
			object returnItem = Invoker.MethodReturn(this, "AddCheckBox", paramsArray);
			NetOffice.PowerPointApi.PPCheckBox newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.PowerPointApi.PPCheckBox.LateBindingApiWrapperType) as NetOffice.PowerPointApi.PPCheckBox;
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
		[SupportByVersionAttribute("PowerPoint", 9)]
		public NetOffice.PowerPointApi.PPRadioCluster AddRadioCluster(Single left, Single top, Single width, Single height)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(left, top, width, height);
			object returnItem = Invoker.MethodReturn(this, "AddRadioCluster", paramsArray);
			NetOffice.PowerPointApi.PPRadioCluster newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.PowerPointApi.PPRadioCluster.LateBindingApiWrapperType) as NetOffice.PowerPointApi.PPRadioCluster;
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
		[SupportByVersionAttribute("PowerPoint", 9)]
		public NetOffice.PowerPointApi.PPStaticText AddStaticText(Single left, Single top, Single width, Single height)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(left, top, width, height);
			object returnItem = Invoker.MethodReturn(this, "AddStaticText", paramsArray);
			NetOffice.PowerPointApi.PPStaticText newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.PowerPointApi.PPStaticText.LateBindingApiWrapperType) as NetOffice.PowerPointApi.PPStaticText;
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
		/// <param name="verticalScrollBar">optional object VerticalScrollBar</param>
		[SupportByVersionAttribute("PowerPoint", 9)]
		public NetOffice.PowerPointApi.PPEditText AddEditText(Single left, Single top, Single width, Single height, object verticalScrollBar)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(left, top, width, height, verticalScrollBar);
			object returnItem = Invoker.MethodReturn(this, "AddEditText", paramsArray);
			NetOffice.PowerPointApi.PPEditText newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.PowerPointApi.PPEditText.LateBindingApiWrapperType) as NetOffice.PowerPointApi.PPEditText;
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
		public NetOffice.PowerPointApi.PPEditText AddEditText(Single left, Single top, Single width, Single height)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(left, top, width, height);
			object returnItem = Invoker.MethodReturn(this, "AddEditText", paramsArray);
			NetOffice.PowerPointApi.PPEditText newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.PowerPointApi.PPEditText.LateBindingApiWrapperType) as NetOffice.PowerPointApi.PPEditText;
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
		[SupportByVersionAttribute("PowerPoint", 9)]
		public NetOffice.PowerPointApi.PPIcon AddIcon(Single left, Single top, Single width, Single height)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(left, top, width, height);
			object returnItem = Invoker.MethodReturn(this, "AddIcon", paramsArray);
			NetOffice.PowerPointApi.PPIcon newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.PowerPointApi.PPIcon.LateBindingApiWrapperType) as NetOffice.PowerPointApi.PPIcon;
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
		[SupportByVersionAttribute("PowerPoint", 9)]
		public NetOffice.PowerPointApi.PPBitmap AddBitmap(Single left, Single top, Single width, Single height)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(left, top, width, height);
			object returnItem = Invoker.MethodReturn(this, "AddBitmap", paramsArray);
			NetOffice.PowerPointApi.PPBitmap newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.PowerPointApi.PPBitmap.LateBindingApiWrapperType) as NetOffice.PowerPointApi.PPBitmap;
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
		[SupportByVersionAttribute("PowerPoint", 9)]
		public NetOffice.PowerPointApi.PPSpinner AddSpinner(Single left, Single top, Single width, Single height)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(left, top, width, height);
			object returnItem = Invoker.MethodReturn(this, "AddSpinner", paramsArray);
			NetOffice.PowerPointApi.PPSpinner newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.PowerPointApi.PPSpinner.LateBindingApiWrapperType) as NetOffice.PowerPointApi.PPSpinner;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9
		/// 
		/// </summary>
		/// <param name="style">NetOffice.PowerPointApi.Enums.PpScrollBarStyle Style</param>
		/// <param name="left">Single Left</param>
		/// <param name="top">Single Top</param>
		/// <param name="width">Single Width</param>
		/// <param name="height">Single Height</param>
		[SupportByVersionAttribute("PowerPoint", 9)]
		public NetOffice.PowerPointApi.PPScrollBar AddScrollBar(NetOffice.PowerPointApi.Enums.PpScrollBarStyle style, Single left, Single top, Single width, Single height)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(style, left, top, width, height);
			object returnItem = Invoker.MethodReturn(this, "AddScrollBar", paramsArray);
			NetOffice.PowerPointApi.PPScrollBar newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.PowerPointApi.PPScrollBar.LateBindingApiWrapperType) as NetOffice.PowerPointApi.PPScrollBar;
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
		[SupportByVersionAttribute("PowerPoint", 9)]
		public NetOffice.PowerPointApi.PPGroupBox AddGroupBox(Single left, Single top, Single width, Single height)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(left, top, width, height);
			object returnItem = Invoker.MethodReturn(this, "AddGroupBox", paramsArray);
			NetOffice.PowerPointApi.PPGroupBox newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.PowerPointApi.PPGroupBox.LateBindingApiWrapperType) as NetOffice.PowerPointApi.PPGroupBox;
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
		[SupportByVersionAttribute("PowerPoint", 9)]
		public NetOffice.PowerPointApi.PPDropDown AddDropDown(Single left, Single top, Single width, Single height)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(left, top, width, height);
			object returnItem = Invoker.MethodReturn(this, "AddDropDown", paramsArray);
			NetOffice.PowerPointApi.PPDropDown newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.PowerPointApi.PPDropDown.LateBindingApiWrapperType) as NetOffice.PowerPointApi.PPDropDown;
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
		[SupportByVersionAttribute("PowerPoint", 9)]
		public NetOffice.PowerPointApi.PPDropDownEdit AddDropDownEdit(Single left, Single top, Single width, Single height)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(left, top, width, height);
			object returnItem = Invoker.MethodReturn(this, "AddDropDownEdit", paramsArray);
			NetOffice.PowerPointApi.PPDropDownEdit newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.PowerPointApi.PPDropDownEdit.LateBindingApiWrapperType) as NetOffice.PowerPointApi.PPDropDownEdit;
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
		[SupportByVersionAttribute("PowerPoint", 9)]
		public NetOffice.PowerPointApi.PPSlideMiniature AddMiniature(Single left, Single top, Single width, Single height)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(left, top, width, height);
			object returnItem = Invoker.MethodReturn(this, "AddMiniature", paramsArray);
			NetOffice.PowerPointApi.PPSlideMiniature newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.PowerPointApi.PPSlideMiniature.LateBindingApiWrapperType) as NetOffice.PowerPointApi.PPSlideMiniature;
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
		[SupportByVersionAttribute("PowerPoint", 9)]
		public NetOffice.PowerPointApi.PPFrame AddFrame(Single left, Single top, Single width, Single height)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(left, top, width, height);
			object returnItem = Invoker.MethodReturn(this, "AddFrame", paramsArray);
			NetOffice.PowerPointApi.PPFrame newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.PowerPointApi.PPFrame.LateBindingApiWrapperType) as NetOffice.PowerPointApi.PPFrame;
			return newObject;
		}

		#endregion
		#pragma warning restore
	}
}