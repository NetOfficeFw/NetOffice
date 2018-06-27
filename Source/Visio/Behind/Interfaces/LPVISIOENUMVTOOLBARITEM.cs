using System.Reflection;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.VisioApi;

namespace NetOffice.VisioApi.Behind
{
	/// <summary>
	/// Interface LPVISIOENUMVTOOLBARITEM 
	/// SupportByVersion Visio, 11,12,14,15,16
	/// </summary>
	[SupportByVersion("Visio", 11,12,14,15,16)]
	[EntityType(EntityType.IsInterface)]
 	public class LPVISIOENUMVTOOLBARITEM : COMObject, NetOffice.VisioApi.LPVISIOENUMVTOOLBARITEM
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
                    _contractType = typeof(NetOffice.VisioApi.LPVISIOENUMVTOOLBARITEM);
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
                    _type = typeof(LPVISIOENUMVTOOLBARITEM);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public LPVISIOENUMVTOOLBARITEM() : base()
		{

		}

		#endregion
		
		#region Properties

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="celt">Int32 celt</param>
		/// <param name="rgelt">NetOffice.VisioApi.IVToolbarItem rgelt</param>
		/// <param name="pceltFetched">Int32 pceltFetched</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual Int32 Next(Int32 celt, out NetOffice.VisioApi.IVToolbarItem rgelt, out Int32 pceltFetched)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false,true,true);
			rgelt = null;
			pceltFetched = 0;
			object[] paramsArray = Invoker.ValidateParamsArray(celt, rgelt, pceltFetched);
			object returnItem = Invoker.MethodReturn(this, "Next", paramsArray);
			rgelt = (NetOffice.VisioApi.IVToolbarItem)paramsArray[1];
			pceltFetched = (Int32)paramsArray[2];
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="celt">Int32 celt</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual Int32 Skip(Int32 celt)
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "Skip", celt);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual Int32 Reset()
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "Reset");
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="ppenm">NetOffice.VisioApi.IEnumVToolbarItem ppenm</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual Int32 Clone(out NetOffice.VisioApi.IEnumVToolbarItem ppenm)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(true);
			ppenm = null;
			object[] paramsArray = Invoker.ValidateParamsArray(ppenm);
			object returnItem = Invoker.MethodReturn(this, "Clone", paramsArray);
			ppenm = (NetOffice.VisioApi.IEnumVToolbarItem)paramsArray[0];
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		#endregion

		#pragma warning restore
	}
}

