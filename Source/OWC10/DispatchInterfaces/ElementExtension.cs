using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.OWC10Api
{
	/// <summary>
	/// DispatchInterface ElementExtension 
	/// SupportByVersion OWC10, 1
	/// </summary>
	[SupportByVersion("OWC10", 1)]
	[EntityType(EntityType.IsDispatchInterface)]
 	public class ElementExtension : COMObject
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
                    _type = typeof(ElementExtension);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public ElementExtension(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public ElementExtension(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ElementExtension(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ElementExtension(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ElementExtension(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ElementExtension(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ElementExtension() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ElementExtension(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public string ElementID
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "ElementID");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ElementID", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public bool ConsumesRecordset
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "ConsumesRecordset");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ConsumesRecordset", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public string AlternateDataSource
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "AlternateDataSource");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "AlternateDataSource", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public string ListRowSource
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "ListRowSource");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ListRowSource", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public string ListBoundField
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "ListBoundField");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ListBoundField", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public string ListDisplayField
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "ListDisplayField");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ListDisplayField", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public string ChildLabel
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "ChildLabel");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ChildLabel", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public NetOffice.OWC10Api.Enums.DscTotalTypeEnum TotalType
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OWC10Api.Enums.DscTotalTypeEnum>(this, "TotalType");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "TotalType", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public string DefaultValue
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "DefaultValue");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "DefaultValue", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
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
		public string ControlSource
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "ControlSource");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ControlSource", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public string UniqueTable
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "UniqueTable");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "UniqueTable", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public string ResyncCommand
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "ResyncCommand");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ResyncCommand", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public string ServerFilter
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "ServerFilter");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ServerFilter", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public string Format
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Format");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Format", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public string RecordsetLabel
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "RecordsetLabel");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "RecordsetLabel", value);
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

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="bstrOldSource">string bstrOldSource</param>
		/// <param name="bstrNewSource">string bstrNewSource</param>
		/// <param name="bstrOldDefaultCaption">string bstrOldDefaultCaption</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("OWC10", 1)]
		public void FixupNames(string bstrOldSource, string bstrNewSource, string bstrOldDefaultCaption)
		{
			 Factory.ExecuteMethod(this, "FixupNames", bstrOldSource, bstrNewSource, bstrOldDefaultCaption);
		}

		#endregion

		#pragma warning restore
	}
}
