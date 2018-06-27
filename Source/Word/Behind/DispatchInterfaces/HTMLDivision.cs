using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.WordApi;

namespace NetOffice.WordApi.Behind
{
	/// <summary>
	/// DispatchInterface HTMLDivision 
	/// SupportByVersion Word, 10,11,12,14,15,16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835683.aspx </remarks>
	[SupportByVersion("Word", 10,11,12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface)]
 	public class HTMLDivision : COMObject, NetOffice.WordApi.HTMLDivision
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
                    _contractType = typeof(NetOffice.WordApi.HTMLDivision);
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
                    _type = typeof(HTMLDivision);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public HTMLDivision() : base()
		{

		}

		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834547.aspx </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.Application Application
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Application>(this, "Application", typeof(NetOffice.WordApi.Application));
			}
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839407.aspx </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public virtual Int32 Creator
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "Creator");
			}
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194751.aspx </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16), ProxyResult]
		public virtual object Parent
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "Parent");
			}
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195903.aspx </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.Range Range
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Range>(this, "Range", typeof(NetOffice.WordApi.Range));
			}
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194413.aspx </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.Borders Borders
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Borders>(this, "Borders", typeof(NetOffice.WordApi.Borders));
			}
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840798.aspx </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public virtual Single LeftIndent
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteSinglePropertyGet(this, "LeftIndent");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "LeftIndent", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838898.aspx </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public virtual Single RightIndent
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteSinglePropertyGet(this, "RightIndent");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "RightIndent", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836995.aspx </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public virtual Single SpaceBefore
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteSinglePropertyGet(this, "SpaceBefore");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "SpaceBefore", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192558.aspx </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public virtual Single SpaceAfter
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteSinglePropertyGet(this, "SpaceAfter");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "SpaceAfter", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff191870.aspx </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.HTMLDivisions HTMLDivisions
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.HTMLDivisions>(this, "HTMLDivisions", typeof(NetOffice.WordApi.HTMLDivisions));
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff841024.aspx </remarks>
		/// <param name="levelsUp">optional object levelsUp</param>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.HTMLDivision HTMLDivisionParent(object levelsUp)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.HTMLDivision>(this, "HTMLDivisionParent", typeof(NetOffice.WordApi.HTMLDivision), levelsUp);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff841024.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.HTMLDivision HTMLDivisionParent()
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.HTMLDivision>(this, "HTMLDivisionParent", typeof(NetOffice.WordApi.HTMLDivision));
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196594.aspx </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public virtual void Delete()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Delete");
		}

		#endregion

		#pragma warning restore
	}
}


