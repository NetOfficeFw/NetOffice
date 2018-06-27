using System.Reflection;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.MSHTMLApi;

namespace NetOffice.MSHTMLApi.Behind
{
	/// <summary>
	/// Interface IHTMLChangeLog 
	/// SupportByVersion MSHTML, 4
	/// </summary>
	[SupportByVersion("MSHTML", 4)]
	[EntityType(EntityType.IsInterface)]
 	public class IHTMLChangeLog : COMObject, NetOffice.MSHTMLApi.IHTMLChangeLog
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
                    _contractType = typeof(NetOffice.MSHTMLApi.IHTMLChangeLog);
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
                    _type = typeof(IHTMLChangeLog);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public IHTMLChangeLog() : base()
		{

		}

		#endregion
		
		#region Properties

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pbBuffer">byte pbBuffer</param>
		/// <param name="nBufferSize">Int32 nBufferSize</param>
		/// <param name="pnRecordLength">Int32 pnRecordLength</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 GetNextChange(byte pbBuffer, Int32 nBufferSize, out Int32 pnRecordLength)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false,false,true);
			pnRecordLength = 0;
			object[] paramsArray = Invoker.ValidateParamsArray(pbBuffer, nBufferSize, pnRecordLength);
			object returnItem = Invoker.MethodReturn(this, "GetNextChange", paramsArray, modifiers);
			pnRecordLength = (Int32)paramsArray[2];
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		#endregion

		#pragma warning restore
	}
}

