using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.OWC10Api;

namespace NetOffice.OWC10Api.Behind
{
	/// <summary>
	/// DispatchInterface Protection 
	/// SupportByVersion OWC10, 1
	/// </summary>
	[SupportByVersion("OWC10", 1)]
	[EntityType(EntityType.IsDispatchInterface)]
 	public class Protection : COMObject, NetOffice.OWC10Api.Protection
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
                    _contractType = typeof(NetOffice.OWC10Api.Protection);
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
                    _type = typeof(Protection);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public Protection() : base()
		{

		}

		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		[BaseResult]
		public virtual NetOffice.OWC10Api.ISpreadsheet Application
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBaseReferencePropertyGet<NetOffice.OWC10Api.ISpreadsheet>(this, "Application");
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public virtual NetOffice.OWC10Api.Worksheet Parent
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OWC10Api.Worksheet>(this, "Parent", typeof(NetOffice.OWC10Api.Worksheet));
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public virtual bool Enabled
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "Enabled");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Enabled", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public virtual bool AllowFormattingRows
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "AllowFormattingRows");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "AllowFormattingRows", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public virtual bool AllowFormattingColumns
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "AllowFormattingColumns");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "AllowFormattingColumns", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public virtual bool AllowInsertingRows
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "AllowInsertingRows");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "AllowInsertingRows", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public virtual bool AllowInsertingColumns
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "AllowInsertingColumns");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "AllowInsertingColumns", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public virtual bool AllowDeletingRows
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "AllowDeletingRows");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "AllowDeletingRows", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public virtual bool AllowDeletingColumns
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "AllowDeletingColumns");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "AllowDeletingColumns", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public virtual bool AllowSorting
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "AllowSorting");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "AllowSorting", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public virtual bool AllowFiltering
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "AllowFiltering");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "AllowFiltering", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public virtual bool AllowHeadingRename
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "AllowHeadingRename");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "AllowHeadingRename", value);
			}
		}

		#endregion

		#region Methods

		#endregion

		#pragma warning restore
	}
}


