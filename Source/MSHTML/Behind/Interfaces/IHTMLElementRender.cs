using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.MSHTMLApi;

namespace NetOffice.MSHTMLApi.Behind
{
	/// <summary>
	/// Interface IHTMLElementRender 
	/// SupportByVersion MSHTML, 4
	/// </summary>
	[SupportByVersion("MSHTML", 4)]
	[EntityType(EntityType.IsInterface)]
 	public class IHTMLElementRender : COMObject, NetOffice.MSHTMLApi.IHTMLElementRender
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
                    _contractType = typeof(NetOffice.MSHTMLApi.IHTMLElementRender);
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
                    _type = typeof(IHTMLElementRender);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public IHTMLElementRender() : base()
		{

		}

		#endregion
		
		#region Properties

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="hdc">_RemotableHandle hdc</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 DrawToDC(_RemotableHandle hdc)
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "DrawToDC", hdc);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="bstrPrinterName">string bstrPrinterName</param>
		/// <param name="hdc">_RemotableHandle hdc</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 SetDocumentPrinter(string bstrPrinterName, _RemotableHandle hdc)
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "SetDocumentPrinter", bstrPrinterName, hdc);
		}

		#endregion

		#pragma warning restore
	}
}

