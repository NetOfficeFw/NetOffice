using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.OWC10Api;

namespace NetOffice.OWC10Api.Behind
{
	/// <summary>
	/// Interface INavUIHost 
	/// SupportByVersion OWC10, 1
	/// </summary>
	[SupportByVersion("OWC10", 1)]
	[EntityType(EntityType.IsInterface)]
 	public class INavUIHost : COMObject, NetOffice.OWC10Api.INavUIHost
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
                    _contractType = typeof(NetOffice.OWC10Api.INavUIHost);
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
                    _type = typeof(INavUIHost);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public INavUIHost() : base()
		{

		}

		#endregion
		
		#region Properties

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="navbtn">Int32 navbtn</param>
		[SupportByVersion("OWC10", 1)]
		public virtual Int32 IsButtonEnabled(Int32 navbtn)
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "IsButtonEnabled", navbtn);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="navbtn">Int32 navbtn</param>
		/// <param name="cancel">Int32 cancel</param>
		[SupportByVersion("OWC10", 1)]
		public virtual Int32 BeforeButtonClick(Int32 navbtn, Int32 cancel)
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "BeforeButtonClick", navbtn, cancel);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="navbtn">Int32 navbtn</param>
		[SupportByVersion("OWC10", 1)]
		public virtual Int32 AfterButtonClick(Int32 navbtn)
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "AfterButtonClick", navbtn);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="displayText">string displayText</param>
		[SupportByVersion("OWC10", 1)]
		public virtual Int32 GetDisplayText(string displayText)
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "GetDisplayText", displayText);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public virtual Int32 OnNavUIChange()
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "OnNavUIChange");
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public virtual Int32 IsFilterOn()
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "IsFilterOn");
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public virtual Int32 IsContextBiDi()
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "IsContextBiDi");
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="fontName">string fontName</param>
		[SupportByVersion("OWC10", 1)]
		public virtual Int32 GetFontName(string fontName)
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "GetFontName", fontName);
		}

		#endregion

		#pragma warning restore
	}
}

