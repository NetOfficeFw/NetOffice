using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.MSHTMLApi;

namespace NetOffice.MSHTMLApi.Behind
{
	/// <summary>
	/// DispatchInterface IHTMLTableSection2 
	/// SupportByVersion MSHTML, 4
	/// </summary>
	[SupportByVersion("MSHTML", 4)]
	[EntityType(EntityType.IsDispatchInterface), BaseType]
 	public class IHTMLTableSection2 : IHTMLTableSection, NetOffice.MSHTMLApi.IHTMLTableSection2
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
                    _contractType = typeof(NetOffice.MSHTMLApi.IHTMLTableSection2);
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
                    _type = typeof(IHTMLTableSection2);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public IHTMLTableSection2() : base()
		{

		}

		#endregion
		
		#region Properties

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="indexFrom">optional Int32 indexFrom = -1</param>
		/// <param name="indexTo">optional Int32 indexTo = -1</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual object moveRow(object indexFrom, object indexTo)
		{
			return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "moveRow", indexFrom, indexTo);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[CustomMethod]
		[SupportByVersion("MSHTML", 4)]
		public virtual object moveRow()
		{
			return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "moveRow");
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="indexFrom">optional Int32 indexFrom = -1</param>
		[CustomMethod]
		[SupportByVersion("MSHTML", 4)]
		public virtual object moveRow(object indexFrom)
		{
			return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "moveRow", indexFrom);
		}

		#endregion

		#pragma warning restore
	}
}

