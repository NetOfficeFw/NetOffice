using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.PowerPointApi
{
	/// <summary>
	/// DispatchInterface PPListBox 
	/// SupportByVersion PowerPoint, 9
	/// </summary>
	[SupportByVersion("PowerPoint", 9)]
	[EntityType(EntityType.IsDispatchInterface)]
 	public class PPListBox : PPControl
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
                    _type = typeof(PPListBox);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public PPListBox(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public PPListBox(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public PPListBox(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public PPListBox(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public PPListBox(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public PPListBox(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public PPListBox() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public PPListBox(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion PowerPoint 9
		/// Get
		/// </summary>
		[SupportByVersion("PowerPoint", 9)]
		public NetOffice.PowerPointApi.PPStrings Strings
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.PowerPointApi.PPStrings>(this, "Strings", NetOffice.PowerPointApi.PPStrings.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9
		/// Get/Set
		/// </summary>
		[SupportByVersion("PowerPoint", 9)]
		public NetOffice.PowerPointApi.Enums.PpListBoxSelectionStyle SelectionStyle
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.PowerPointApi.Enums.PpListBoxSelectionStyle>(this, "SelectionStyle");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "SelectionStyle", value);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9
		/// Get/Set
		/// </summary>
		[SupportByVersion("PowerPoint", 9)]
		public Int32 FocusItem
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "FocusItem");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "FocusItem", value);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9
		/// Get
		/// </summary>
		[SupportByVersion("PowerPoint", 9)]
		public Int32 TopItem
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "TopItem");
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9
		/// Get/Set
		/// </summary>
		[SupportByVersion("PowerPoint", 9)]
		public string OnSelectionChange
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "OnSelectionChange");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "OnSelectionChange", value);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9
		/// Get/Set
		/// </summary>
		[SupportByVersion("PowerPoint", 9)]
		public string OnDoubleClick
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "OnDoubleClick");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "OnDoubleClick", value);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9
		/// Get/Set
		/// </summary>
		/// <param name="index">Int32 index</param>
		[SupportByVersion("PowerPoint", 9)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NetOffice.OfficeApi.Enums.MsoTriState get_IsSelected(Int32 index)
		{
			return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoTriState>(this, "IsSelected", index);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9
		/// Get/Set
		/// </summary>
		/// <param name="index">Int32 index</param>
		[SupportByVersion("PowerPoint", 9)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public void set_IsSelected(Int32 index, NetOffice.OfficeApi.Enums.MsoTriState value)
		{
			Factory.ExecutePropertySet(this, "IsSelected", index, value);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9
		/// Alias for get_IsSelected
		/// </summary>
		/// <param name="index">Int32 index</param>
		[SupportByVersion("PowerPoint", 9), Redirect("get_IsSelected")]
		public NetOffice.OfficeApi.Enums.MsoTriState IsSelected(Int32 index)
		{
			return get_IsSelected(index);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9
		/// Get
		/// </summary>
		[SupportByVersion("PowerPoint", 9)]
		public NetOffice.PowerPointApi.Enums.PpListBoxAbbreviationStyle IsAbbreviated
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.PowerPointApi.Enums.PpListBoxAbbreviationStyle>(this, "IsAbbreviated");
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion PowerPoint 9
		/// </summary>
		/// <param name="safeArrayTabStops">object safeArrayTabStops</param>
		[SupportByVersion("PowerPoint", 9)]
		public void SetTabStops(object safeArrayTabStops)
		{
			 Factory.ExecuteMethod(this, "SetTabStops", safeArrayTabStops);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9
		/// </summary>
		/// <param name="style">NetOffice.PowerPointApi.Enums.PpListBoxAbbreviationStyle style</param>
		[SupportByVersion("PowerPoint", 9)]
		public void Abbreviate(NetOffice.PowerPointApi.Enums.PpListBoxAbbreviationStyle style)
		{
			 Factory.ExecuteMethod(this, "Abbreviate", style);
		}

		#endregion

		#pragma warning restore
	}
}
