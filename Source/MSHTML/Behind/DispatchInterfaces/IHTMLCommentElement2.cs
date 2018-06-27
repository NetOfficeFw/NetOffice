using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.MSHTMLApi;

namespace NetOffice.MSHTMLApi.Behind
{
	/// <summary>
	/// DispatchInterface IHTMLCommentElement2 
	/// SupportByVersion MSHTML, 4
	/// </summary>
	[SupportByVersion("MSHTML", 4)]
	[EntityType(EntityType.IsDispatchInterface)]
 	public class IHTMLCommentElement2 : IHTMLCommentElement, NetOffice.MSHTMLApi.IHTMLCommentElement2
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
                    _contractType = typeof(NetOffice.MSHTMLApi.IHTMLCommentElement2);
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
                    _type = typeof(IHTMLCommentElement2);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public IHTMLCommentElement2() : base()
		{

		}

		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual string data
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "data");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "data", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// Get
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 length
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "length");
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="offset">Int32 offset</param>
		/// <param name="count">Int32 count</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual string substringData(Int32 offset, Int32 count)
		{
			return InvokerService.InvokeInternal.ExecuteStringMethodGet(this, "substringData", offset, count);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="bstrstring">string bstrstring</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual void appendData(string bstrstring)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "appendData", bstrstring);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="offset">Int32 offset</param>
		/// <param name="bstrstring">string bstrstring</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual void insertData(Int32 offset, string bstrstring)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "insertData", offset, bstrstring);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="offset">Int32 offset</param>
		/// <param name="count">Int32 count</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual void deleteData(Int32 offset, Int32 count)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "deleteData", offset, count);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="offset">Int32 offset</param>
		/// <param name="count">Int32 count</param>
		/// <param name="bstrstring">string bstrstring</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual void replaceData(Int32 offset, Int32 count, string bstrstring)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "replaceData", offset, count, bstrstring);
		}

		#endregion

		#pragma warning restore
	}
}

