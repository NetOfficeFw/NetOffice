using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.OWC10Api;

namespace NetOffice.OWC10Api.Behind
{
	/// <summary>
	/// DispatchInterface ElementExtension 
	/// SupportByVersion OWC10, 1
	/// </summary>
	[SupportByVersion("OWC10", 1)]
	[EntityType(EntityType.IsDispatchInterface)]
 	public class ElementExtension : COMObject, NetOffice.OWC10Api.ElementExtension
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
                    _contractType = typeof(NetOffice.OWC10Api.ElementExtension);
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
                    _type = typeof(ElementExtension);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public ElementExtension() : base()
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
		public virtual string ElementID
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "ElementID");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ElementID", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public virtual bool ConsumesRecordset
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "ConsumesRecordset");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ConsumesRecordset", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public virtual string AlternateDataSource
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "AlternateDataSource");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "AlternateDataSource", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public virtual string ListRowSource
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "ListRowSource");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ListRowSource", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public virtual string ListBoundField
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "ListBoundField");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ListBoundField", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public virtual string ListDisplayField
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "ListDisplayField");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ListDisplayField", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual string ChildLabel
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "ChildLabel");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ChildLabel", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public virtual NetOffice.OWC10Api.Enums.DscTotalTypeEnum TotalType
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.OWC10Api.Enums.DscTotalTypeEnum>(this, "TotalType");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "TotalType", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public virtual string DefaultValue
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "DefaultValue");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "DefaultValue", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public virtual string RecordSource
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "RecordSource");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "RecordSource", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public virtual string ControlSource
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "ControlSource");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ControlSource", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public virtual string UniqueTable
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "UniqueTable");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "UniqueTable", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public virtual string ResyncCommand
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "ResyncCommand");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ResyncCommand", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public virtual string ServerFilter
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "ServerFilter");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ServerFilter", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public virtual string Format
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "Format");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Format", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public virtual string RecordsetLabel
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "RecordsetLabel");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "RecordsetLabel", value);
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
		public virtual void SetDesignerFlags(Int32 grfFlags, bool vfSet)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SetDesignerFlags", grfFlags, vfSet);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="bstrOldSource">string bstrOldSource</param>
		/// <param name="bstrNewSource">string bstrNewSource</param>
		/// <param name="bstrOldDefaultCaption">string bstrOldDefaultCaption</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("OWC10", 1)]
		public virtual void FixupNames(string bstrOldSource, string bstrNewSource, string bstrOldDefaultCaption)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "FixupNames", bstrOldSource, bstrNewSource, bstrOldDefaultCaption);
		}

		#endregion

		#pragma warning restore
	}
}

