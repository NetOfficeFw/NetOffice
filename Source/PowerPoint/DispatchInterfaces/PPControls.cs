using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.PowerPointApi
{
	/// <summary>
	/// DispatchInterface PPControls 
	/// SupportByVersion PowerPoint, 9
	/// </summary>
	[SupportByVersion("PowerPoint", 9)]
	[EntityType(EntityType.IsDispatchInterface), HasIndexProperty(IndexInvoke.Method, "Item")]
 	public class PPControls : Collection
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
                    _type = typeof(PPControls);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public PPControls(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

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
		
		/// <param name="progId">registered progID</param>
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
		/// Get/Set
		/// </summary>
		[SupportByVersion("PowerPoint", 9)]
		public NetOffice.OfficeApi.Enums.MsoTriState Visible
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoTriState>(this, "Visible");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "Visible", value);
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion PowerPoint 9
		/// </summary>
		/// <param name="index">object index</param>
		[SupportByVersion("PowerPoint", 9)]
		[BaseResult]
		[NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item"), IndexProperty]
		public NetOffice.PowerPointApi.PPControl this[object index]
		{
			get
			{
				return Factory.ExecuteBaseReferenceMethodGet<NetOffice.PowerPointApi.PPControl>(this, "Item", index);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9
		/// </summary>
		/// <param name="left">Single left</param>
		/// <param name="top">Single top</param>
		/// <param name="width">Single width</param>
		/// <param name="height">Single height</param>
		[SupportByVersion("PowerPoint", 9)]
		public NetOffice.PowerPointApi.PPPushButton AddPushButton(Single left, Single top, Single width, Single height)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.PPPushButton>(this, "AddPushButton", NetOffice.PowerPointApi.PPPushButton.LateBindingApiWrapperType, left, top, width, height);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9
		/// </summary>
		/// <param name="left">Single left</param>
		/// <param name="top">Single top</param>
		/// <param name="width">Single width</param>
		/// <param name="height">Single height</param>
		[SupportByVersion("PowerPoint", 9)]
		public NetOffice.PowerPointApi.PPToggleButton AddToggleButton(Single left, Single top, Single width, Single height)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.PPToggleButton>(this, "AddToggleButton", NetOffice.PowerPointApi.PPToggleButton.LateBindingApiWrapperType, left, top, width, height);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9
		/// </summary>
		/// <param name="left">Single left</param>
		/// <param name="top">Single top</param>
		/// <param name="width">Single width</param>
		/// <param name="height">Single height</param>
		[SupportByVersion("PowerPoint", 9)]
		public NetOffice.PowerPointApi.PPBitmapButton AddBitmapButton(Single left, Single top, Single width, Single height)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.PPBitmapButton>(this, "AddBitmapButton", NetOffice.PowerPointApi.PPBitmapButton.LateBindingApiWrapperType, left, top, width, height);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9
		/// </summary>
		/// <param name="left">Single left</param>
		/// <param name="top">Single top</param>
		/// <param name="width">Single width</param>
		/// <param name="height">Single height</param>
		[SupportByVersion("PowerPoint", 9)]
		public NetOffice.PowerPointApi.PPListBox AddListBox(Single left, Single top, Single width, Single height)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.PPListBox>(this, "AddListBox", NetOffice.PowerPointApi.PPListBox.LateBindingApiWrapperType, left, top, width, height);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9
		/// </summary>
		/// <param name="left">Single left</param>
		/// <param name="top">Single top</param>
		/// <param name="width">Single width</param>
		/// <param name="height">Single height</param>
		[SupportByVersion("PowerPoint", 9)]
		public NetOffice.PowerPointApi.PPCheckBox AddCheckBox(Single left, Single top, Single width, Single height)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.PPCheckBox>(this, "AddCheckBox", NetOffice.PowerPointApi.PPCheckBox.LateBindingApiWrapperType, left, top, width, height);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9
		/// </summary>
		/// <param name="left">Single left</param>
		/// <param name="top">Single top</param>
		/// <param name="width">Single width</param>
		/// <param name="height">Single height</param>
		[SupportByVersion("PowerPoint", 9)]
		public NetOffice.PowerPointApi.PPRadioCluster AddRadioCluster(Single left, Single top, Single width, Single height)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.PPRadioCluster>(this, "AddRadioCluster", NetOffice.PowerPointApi.PPRadioCluster.LateBindingApiWrapperType, left, top, width, height);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9
		/// </summary>
		/// <param name="left">Single left</param>
		/// <param name="top">Single top</param>
		/// <param name="width">Single width</param>
		/// <param name="height">Single height</param>
		[SupportByVersion("PowerPoint", 9)]
		public NetOffice.PowerPointApi.PPStaticText AddStaticText(Single left, Single top, Single width, Single height)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.PPStaticText>(this, "AddStaticText", NetOffice.PowerPointApi.PPStaticText.LateBindingApiWrapperType, left, top, width, height);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9
		/// </summary>
		/// <param name="left">Single left</param>
		/// <param name="top">Single top</param>
		/// <param name="width">Single width</param>
		/// <param name="height">Single height</param>
		/// <param name="verticalScrollBar">optional object verticalScrollBar</param>
		[SupportByVersion("PowerPoint", 9)]
		public NetOffice.PowerPointApi.PPEditText AddEditText(Single left, Single top, Single width, Single height, object verticalScrollBar)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.PPEditText>(this, "AddEditText", NetOffice.PowerPointApi.PPEditText.LateBindingApiWrapperType, new object[]{ left, top, width, height, verticalScrollBar });
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
		public NetOffice.PowerPointApi.PPEditText AddEditText(Single left, Single top, Single width, Single height)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.PPEditText>(this, "AddEditText", NetOffice.PowerPointApi.PPEditText.LateBindingApiWrapperType, left, top, width, height);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9
		/// </summary>
		/// <param name="left">Single left</param>
		/// <param name="top">Single top</param>
		/// <param name="width">Single width</param>
		/// <param name="height">Single height</param>
		[SupportByVersion("PowerPoint", 9)]
		public NetOffice.PowerPointApi.PPIcon AddIcon(Single left, Single top, Single width, Single height)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.PPIcon>(this, "AddIcon", NetOffice.PowerPointApi.PPIcon.LateBindingApiWrapperType, left, top, width, height);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9
		/// </summary>
		/// <param name="left">Single left</param>
		/// <param name="top">Single top</param>
		/// <param name="width">Single width</param>
		/// <param name="height">Single height</param>
		[SupportByVersion("PowerPoint", 9)]
		public NetOffice.PowerPointApi.PPBitmap AddBitmap(Single left, Single top, Single width, Single height)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.PPBitmap>(this, "AddBitmap", NetOffice.PowerPointApi.PPBitmap.LateBindingApiWrapperType, left, top, width, height);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9
		/// </summary>
		/// <param name="left">Single left</param>
		/// <param name="top">Single top</param>
		/// <param name="width">Single width</param>
		/// <param name="height">Single height</param>
		[SupportByVersion("PowerPoint", 9)]
		public NetOffice.PowerPointApi.PPSpinner AddSpinner(Single left, Single top, Single width, Single height)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.PPSpinner>(this, "AddSpinner", NetOffice.PowerPointApi.PPSpinner.LateBindingApiWrapperType, left, top, width, height);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9
		/// </summary>
		/// <param name="style">NetOffice.PowerPointApi.Enums.PpScrollBarStyle style</param>
		/// <param name="left">Single left</param>
		/// <param name="top">Single top</param>
		/// <param name="width">Single width</param>
		/// <param name="height">Single height</param>
		[SupportByVersion("PowerPoint", 9)]
		public NetOffice.PowerPointApi.PPScrollBar AddScrollBar(NetOffice.PowerPointApi.Enums.PpScrollBarStyle style, Single left, Single top, Single width, Single height)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.PPScrollBar>(this, "AddScrollBar", NetOffice.PowerPointApi.PPScrollBar.LateBindingApiWrapperType, new object[]{ style, left, top, width, height });
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9
		/// </summary>
		/// <param name="left">Single left</param>
		/// <param name="top">Single top</param>
		/// <param name="width">Single width</param>
		/// <param name="height">Single height</param>
		[SupportByVersion("PowerPoint", 9)]
		public NetOffice.PowerPointApi.PPGroupBox AddGroupBox(Single left, Single top, Single width, Single height)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.PPGroupBox>(this, "AddGroupBox", NetOffice.PowerPointApi.PPGroupBox.LateBindingApiWrapperType, left, top, width, height);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9
		/// </summary>
		/// <param name="left">Single left</param>
		/// <param name="top">Single top</param>
		/// <param name="width">Single width</param>
		/// <param name="height">Single height</param>
		[SupportByVersion("PowerPoint", 9)]
		public NetOffice.PowerPointApi.PPDropDown AddDropDown(Single left, Single top, Single width, Single height)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.PPDropDown>(this, "AddDropDown", NetOffice.PowerPointApi.PPDropDown.LateBindingApiWrapperType, left, top, width, height);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9
		/// </summary>
		/// <param name="left">Single left</param>
		/// <param name="top">Single top</param>
		/// <param name="width">Single width</param>
		/// <param name="height">Single height</param>
		[SupportByVersion("PowerPoint", 9)]
		public NetOffice.PowerPointApi.PPDropDownEdit AddDropDownEdit(Single left, Single top, Single width, Single height)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.PPDropDownEdit>(this, "AddDropDownEdit", NetOffice.PowerPointApi.PPDropDownEdit.LateBindingApiWrapperType, left, top, width, height);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9
		/// </summary>
		/// <param name="left">Single left</param>
		/// <param name="top">Single top</param>
		/// <param name="width">Single width</param>
		/// <param name="height">Single height</param>
		[SupportByVersion("PowerPoint", 9)]
		public NetOffice.PowerPointApi.PPSlideMiniature AddMiniature(Single left, Single top, Single width, Single height)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.PPSlideMiniature>(this, "AddMiniature", NetOffice.PowerPointApi.PPSlideMiniature.LateBindingApiWrapperType, left, top, width, height);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9
		/// </summary>
		/// <param name="left">Single left</param>
		/// <param name="top">Single top</param>
		/// <param name="width">Single width</param>
		/// <param name="height">Single height</param>
		[SupportByVersion("PowerPoint", 9)]
		public NetOffice.PowerPointApi.PPFrame AddFrame(Single left, Single top, Single width, Single height)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.PPFrame>(this, "AddFrame", NetOffice.PowerPointApi.PPFrame.LateBindingApiWrapperType, left, top, width, height);
		}

		#endregion

		#pragma warning restore
	}
}
