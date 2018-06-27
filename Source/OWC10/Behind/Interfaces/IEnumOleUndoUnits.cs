using System.Reflection;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.OWC10Api;

namespace NetOffice.OWC10Api.Behind
{
	/// <summary>
	/// Interface IEnumOleUndoUnits 
	/// SupportByVersion OWC10, 1
	/// </summary>
	[SupportByVersion("OWC10", 1)]
	[EntityType(EntityType.IsInterface)]
 	public class IEnumOleUndoUnits : COMObject, NetOffice.OWC10Api.IEnumOleUndoUnits
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
                    _contractType = typeof(NetOffice.OWC10Api.IEnumOleUndoUnits);
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
                    _type = typeof(IEnumOleUndoUnits);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public IEnumOleUndoUnits() : base()
		{

		}

		#endregion
		
		#region Properties

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="cElt">Int32 cElt</param>
		/// <param name="rgElt">NetOffice.OWC10Api.IOleUndoUnit rgElt</param>
		/// <param name="pcEltFetched">Int32 pcEltFetched</param>
		[SupportByVersion("OWC10", 1)]
		public virtual Int32 RemoteNext(Int32 cElt, out NetOffice.OWC10Api.IOleUndoUnit rgElt, out Int32 pcEltFetched)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false,true,true);
			rgElt = null;
			pcEltFetched = 0;
			object[] paramsArray = Invoker.ValidateParamsArray(cElt, rgElt, pcEltFetched);
			object returnItem = Invoker.MethodReturn(this, "RemoteNext", paramsArray, modifiers);
            if (paramsArray[1] is MarshalByRefObject)
                rgElt = Factory.CreateObjectFromComProxy(this, paramsArray[1], false) as NetOffice.OWC10Api.IOleUndoUnit;
            else
                rgElt = null;
			pcEltFetched = (Int32)paramsArray[2];
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
        }

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="cElt">Int32 cElt</param>
		[SupportByVersion("OWC10", 1)]
		public virtual Int32 Skip(Int32 cElt)
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "Skip", cElt);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public virtual Int32 Reset()
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "Reset");
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="ppEnum">NetOffice.OWC10Api.IEnumOleUndoUnits ppEnum</param>
		[SupportByVersion("OWC10", 1)]
		public virtual Int32 Clone(out NetOffice.OWC10Api.IEnumOleUndoUnits ppEnum)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(true);
			ppEnum = null;
			object[] paramsArray = Invoker.ValidateParamsArray(ppEnum);
			object returnItem = Invoker.MethodReturn(this, "Clone", paramsArray, modifiers);
            if (paramsArray[0] is MarshalByRefObject)
                ppEnum = Factory.CreateObjectFromComProxy(this, paramsArray[0], false) as NetOffice.OWC10Api.IEnumOleUndoUnits;
            else
                ppEnum = null;
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		#endregion

		#pragma warning restore
	}
}

