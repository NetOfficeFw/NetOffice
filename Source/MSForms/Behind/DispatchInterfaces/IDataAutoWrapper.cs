using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.MSFormsApi;

namespace NetOffice.MSFormsApi.Behind
{
	/// <summary>
	/// DispatchInterface IDataAutoWrapper 
	/// SupportByVersion MSForms, 2
	/// </summary>
	[SupportByVersion("MSForms", 2)]
	[EntityType(EntityType.IsDispatchInterface), BaseType]
 	public class IDataAutoWrapper : COMObject, NetOffice.MSFormsApi.IDataAutoWrapper
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
                    _contractType = typeof(NetOffice.MSFormsApi.IDataAutoWrapper);
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
                    _type = typeof(IDataAutoWrapper);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public IDataAutoWrapper() : base()
		{

		}

		#endregion
		
		#region Properties

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		public virtual void Clear()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Clear");
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		/// <param name="format">object format</param>
		[SupportByVersion("MSForms", 2)]
		public virtual bool GetFormat(object format)
		{
			return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "GetFormat", format);
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		/// <param name="format">optional object format</param>
		[SupportByVersion("MSForms", 2)]
		public virtual string GetText(object format)
		{
			return InvokerService.InvokeInternal.ExecuteStringMethodGet(this, "GetText", format);
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		[CustomMethod]
		[SupportByVersion("MSForms", 2)]
		public virtual string GetText()
		{
			return InvokerService.InvokeInternal.ExecuteStringMethodGet(this, "GetText");
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		/// <param name="text">string text</param>
		/// <param name="format">optional object format</param>
		[SupportByVersion("MSForms", 2)]
		public virtual void SetText(string text, object format)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SetText", text, format);
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		/// <param name="text">string text</param>
		[CustomMethod]
		[SupportByVersion("MSForms", 2)]
		public virtual void SetText(string text)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SetText", text);
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		public virtual void PutInClipboard()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "PutInClipboard");
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		public virtual void GetFromClipboard()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "GetFromClipboard");
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		/// <param name="oKEffect">optional object oKEffect</param>
		[SupportByVersion("MSForms", 2)]
		public virtual NetOffice.MSFormsApi.Enums.fmDropEffect StartDrag(object oKEffect)
		{
			return InvokerService.InvokeInternal.ExecuteEnumMethodGet<NetOffice.MSFormsApi.Enums.fmDropEffect>(this, "StartDrag", oKEffect);
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		[CustomMethod]
		[SupportByVersion("MSForms", 2)]
		public virtual NetOffice.MSFormsApi.Enums.fmDropEffect StartDrag()
		{
			return InvokerService.InvokeInternal.ExecuteEnumMethodGet<NetOffice.MSFormsApi.Enums.fmDropEffect>(this, "StartDrag");
		}

		#endregion

		#pragma warning restore
	}
}

