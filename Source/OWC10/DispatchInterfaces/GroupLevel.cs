using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.OWC10Api
{
	/// <summary>
	/// DispatchInterface GroupLevel 
	/// SupportByVersion OWC10, 1
	/// </summary>
	[SupportByVersion("OWC10", 1)]
	[EntityType(EntityType.IsDispatchInterface)]
 	public class GroupLevel : COMObject
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
                    _type = typeof(GroupLevel);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public GroupLevel(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public GroupLevel(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public GroupLevel(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public GroupLevel(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public GroupLevel(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public GroupLevel(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public GroupLevel() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public GroupLevel(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public NetOffice.OWC10Api.Enums.DscGroupOnEnum GroupOn
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OWC10Api.Enums.DscGroupOnEnum>(this, "GroupOn");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "GroupOn", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public Double GroupInterval
		{
			get
			{
				return Factory.ExecuteDoublePropertyGet(this, "GroupInterval");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "GroupInterval", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public bool GroupHeader
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "GroupHeader");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "GroupHeader", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public bool GroupFooter
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "GroupFooter");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "GroupFooter", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public bool CaptionSection
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "CaptionSection");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "CaptionSection", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public bool RecordNavigationSection
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "RecordNavigationSection");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "RecordNavigationSection", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public Int32 DataPageSize
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "DataPageSize");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "DataPageSize", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public bool ExpandedByDefault
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "ExpandedByDefault");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ExpandedByDefault", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public string GroupFilterControl
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "GroupFilterControl");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "GroupFilterControl", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public string DefaultSort
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "DefaultSort");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "DefaultSort", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public string RecordSource
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "RecordSource");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "RecordSource", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public string CaptionElementId
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "CaptionElementId");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "CaptionElementId", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public string HeaderElementId
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "HeaderElementId");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "HeaderElementId", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public string FooterElementId
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "FooterElementId");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "FooterElementId", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public string RecordNavigationElementId
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "RecordNavigationElementId");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "RecordNavigationElementId", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NetOffice.OWC10Api.PageField GroupedOnField
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OWC10Api.PageField>(this, "GroupedOnField", NetOffice.OWC10Api.PageField.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public string GroupFilterField
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "GroupFilterField");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "GroupFilterField", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Int32 SGWindow
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "SGWindow");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "SGWindow", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public UIntPtr SGMessage
		{
			get
			{
				return Factory.ExecuteUIntPtrPropertyGet(this, "SGMessage");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "SGMessage", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public bool AllowEdits
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "AllowEdits");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "AllowEdits", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public bool AllowAdditions
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "AllowAdditions");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "AllowAdditions", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public bool AllowDeletions
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "AllowDeletions");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "AllowDeletions", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public bool RecordSelector
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "RecordSelector");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "RecordSelector", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public string AlternateRowColor
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "AlternateRowColor");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "AlternateRowColor", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public string Name
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Name");
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="grfFlags">Int32 grfFlags</param>
		/// <param name="vfSet">bool vfSet</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("OWC10", 1)]
		public void SetDesignerFlags(Int32 grfFlags, bool vfSet)
		{
			 Factory.ExecuteMethod(this, "SetDesignerFlags", grfFlags, vfSet);
		}

		#endregion

		#pragma warning restore
	}
}
