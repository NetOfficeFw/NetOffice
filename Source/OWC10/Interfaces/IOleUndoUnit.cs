using System.Reflection;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.OWC10Api
{
	/// <summary>
	/// Interface IOleUndoUnit 
	/// SupportByVersion OWC10, 1
	/// </summary>
	[SupportByVersion("OWC10", 1)]
	[EntityType(EntityType.IsInterface), BaseType]
 	public class IOleUndoUnit : COMObject
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
                    _type = typeof(IOleUndoUnit);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public IOleUndoUnit(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public IOleUndoUnit(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IOleUndoUnit(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IOleUndoUnit(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IOleUndoUnit(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IOleUndoUnit(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IOleUndoUnit() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IOleUndoUnit(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="pUndoManager">NetOffice.OWC10Api.IOleUndoManager pUndoManager</param>
		[SupportByVersion("OWC10", 1)]
		public Int32 Do(NetOffice.OWC10Api.IOleUndoManager pUndoManager)
		{
			return Factory.ExecuteInt32MethodGet(this, "Do", pUndoManager);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="pbstr">string pbstr</param>
		[SupportByVersion("OWC10", 1)]
		public Int32 GetDescription(out string pbstr)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(true);
			pbstr = string.Empty;
			object[] paramsArray = Invoker.ValidateParamsArray(pbstr);
			object returnItem = Invoker.MethodReturn(this, "GetDescription", paramsArray, modifiers);
			pbstr = paramsArray[0] as string;
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="pClsid">Guid pClsid</param>
		/// <param name="plID">Int32 plID</param>
		[SupportByVersion("OWC10", 1)]
		public Int32 GetUnitType(out Guid pClsid, out Int32 plID)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(true,true);
			pClsid = Guid.Empty;
			plID = 0;
			object[] paramsArray = Invoker.ValidateParamsArray(pClsid, plID);
			object returnItem = Invoker.MethodReturn(this, "GetUnitType", paramsArray, modifiers);
			pClsid = (Guid)paramsArray[0];
			plID = (Int32)paramsArray[1];
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public Int32 OnNextAdd()
		{
			return Factory.ExecuteInt32MethodGet(this, "OnNextAdd");
		}

		#endregion

		#pragma warning restore
	}
}
