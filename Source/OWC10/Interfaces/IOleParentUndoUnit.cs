using System.Reflection;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.OWC10Api
{
	/// <summary>
	/// Interface IOleParentUndoUnit 
	/// SupportByVersion OWC10, 1
	/// </summary>
	[SupportByVersion("OWC10", 1)]
	[EntityType(EntityType.IsInterface)]
 	public class IOleParentUndoUnit : IOleUndoUnit
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
                    _type = typeof(IOleParentUndoUnit);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public IOleParentUndoUnit(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public IOleParentUndoUnit(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IOleParentUndoUnit(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IOleParentUndoUnit(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IOleParentUndoUnit(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IOleParentUndoUnit(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IOleParentUndoUnit() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IOleParentUndoUnit(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="pPUU">NetOffice.OWC10Api.IOleParentUndoUnit pPUU</param>
		[SupportByVersion("OWC10", 1)]
		public Int32 Open(NetOffice.OWC10Api.IOleParentUndoUnit pPUU)
		{
			return Factory.ExecuteInt32MethodGet(this, "Open", pPUU);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="pPUU">NetOffice.OWC10Api.IOleParentUndoUnit pPUU</param>
		/// <param name="fCommit">Int32 fCommit</param>
		[SupportByVersion("OWC10", 1)]
		public Int32 Close(NetOffice.OWC10Api.IOleParentUndoUnit pPUU, Int32 fCommit)
		{
			return Factory.ExecuteInt32MethodGet(this, "Close", pPUU, fCommit);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="pUU">NetOffice.OWC10Api.IOleUndoUnit pUU</param>
		[SupportByVersion("OWC10", 1)]
		public Int32 Add(NetOffice.OWC10Api.IOleUndoUnit pUU)
		{
			return Factory.ExecuteInt32MethodGet(this, "Add", pUU);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="pUU">NetOffice.OWC10Api.IOleUndoUnit pUU</param>
		[SupportByVersion("OWC10", 1)]
		public Int32 FindUnit(NetOffice.OWC10Api.IOleUndoUnit pUU)
		{
			return Factory.ExecuteInt32MethodGet(this, "FindUnit", pUU);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="pdwState">Int32 pdwState</param>
		[SupportByVersion("OWC10", 1)]
		public Int32 GetParentState(out Int32 pdwState)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(true);
			pdwState = 0;
			object[] paramsArray = Invoker.ValidateParamsArray(pdwState);
			object returnItem = Invoker.MethodReturn(this, "GetParentState", paramsArray, modifiers);
			pdwState = (Int32)paramsArray[0];
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		#endregion

		#pragma warning restore
	}
}
