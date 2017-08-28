using System.Reflection;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.OWC10Api
{
	/// <summary>
	/// Interface IEnumOleUndoUnits 
	/// SupportByVersion OWC10, 1
	/// </summary>
	[SupportByVersion("OWC10", 1)]
	[EntityType(EntityType.IsInterface)]
 	public class IEnumOleUndoUnits : COMObject
	{
		#pragma warning disable

		#region Type Information

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

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public IEnumOleUndoUnits(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public IEnumOleUndoUnits(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IEnumOleUndoUnits(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IEnumOleUndoUnits(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IEnumOleUndoUnits(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IEnumOleUndoUnits(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IEnumOleUndoUnits() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IEnumOleUndoUnits(string progId) : base(progId)
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
		public Int32 RemoteNext(Int32 cElt, out NetOffice.OWC10Api.IOleUndoUnit rgElt, out Int32 pcEltFetched)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false,true,true);
			rgElt = null;
			pcEltFetched = 0;
			object[] paramsArray = Invoker.ValidateParamsArray(cElt, rgElt, pcEltFetched);
			object returnItem = Invoker.MethodReturn(this, "RemoteNext", paramsArray, modifiers);
            if (paramsArray[1] is MarshalByRefObject)
                rgElt = new NetOffice.OWC10Api.IOleUndoUnit(this, paramsArray[1]);
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
		public Int32 Skip(Int32 cElt)
		{
			return Factory.ExecuteInt32MethodGet(this, "Skip", cElt);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public Int32 Reset()
		{
			return Factory.ExecuteInt32MethodGet(this, "Reset");
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="ppEnum">NetOffice.OWC10Api.IEnumOleUndoUnits ppEnum</param>
		[SupportByVersion("OWC10", 1)]
		public Int32 Clone(out NetOffice.OWC10Api.IEnumOleUndoUnits ppEnum)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(true);
			ppEnum = null;
			object[] paramsArray = Invoker.ValidateParamsArray(ppEnum);
			object returnItem = Invoker.MethodReturn(this, "Clone", paramsArray, modifiers);
            if (paramsArray[0] is MarshalByRefObject)
                ppEnum = new NetOffice.OWC10Api.IEnumOleUndoUnits(this, paramsArray[0]);
            else
                ppEnum = null;
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		#endregion

		#pragma warning restore
	}
}
