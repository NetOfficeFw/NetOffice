using System.Reflection;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.WordApi;

namespace NetOffice.WordApi.Behind
{
	/// <summary>
	/// DispatchInterface ConnectorFormat 
	/// SupportByVersion Word, 9,10,11,12,14,15,16
	/// </summary>
	[SupportByVersion("Word", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface)]
 	public class ConnectorFormat : COMObject, NetOffice.WordApi.ConnectorFormat
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
                    _contractType = typeof(NetOffice.WordApi.ConnectorFormat);
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
                    _type = typeof(ConnectorFormat);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public ConnectorFormat() : base()
		{

		}

		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.Application Application
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Application>(this, "Application", typeof(NetOffice.WordApi.Application));
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual Int32 Creator
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "Creator");
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.OfficeApi.Enums.MsoTriState BeginConnected
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoTriState>(this, "BeginConnected");
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.Shape BeginConnectedShape
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Shape>(this, "BeginConnectedShape", typeof(NetOffice.WordApi.Shape));
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual Int32 BeginConnectionSite
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "BeginConnectionSite");
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.OfficeApi.Enums.MsoTriState EndConnected
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoTriState>(this, "EndConnected");
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.Shape EndConnectedShape
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Shape>(this, "EndConnectedShape", typeof(NetOffice.WordApi.Shape));
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual Int32 EndConnectionSite
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "EndConnectionSite");
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("Word", 9,10,11,12,14,15,16), ProxyResult]
		public virtual object Parent
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "Parent");
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.OfficeApi.Enums.MsoConnectorType Type
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoConnectorType>(this, "Type");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "Type", value);
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="connectedShape">NetOffice.WordApi.Shape connectedShape</param>
		/// <param name="connectionSite">Int32 connectionSite</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void BeginConnect(out NetOffice.WordApi.Shape connectedShape, Int32 connectionSite)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(true,false);
			connectedShape = null;
			object[] paramsArray = Invoker.ValidateParamsArray(connectedShape, connectionSite);
            InvokerService.InvokeInternal.ExecuteMethod(this, "BeginConnect", paramsArray, modifiers); 
            if (paramsArray[0] is MarshalByRefObject)
                connectedShape = Factory.CreateKnownObjectFromComProxy<NetOffice.WordApi.Shape>(this, paramsArray[0], typeof(NetOffice.WordApi.Shape));
            else
                connectedShape = null;
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void BeginDisconnect()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "BeginDisconnect");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="connectedShape">NetOffice.WordApi.Shape connectedShape</param>
		/// <param name="connectionSite">Int32 connectionSite</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void EndConnect(out NetOffice.WordApi.Shape connectedShape, Int32 connectionSite)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(true,false);
			connectedShape = null;
			object[] paramsArray = Invoker.ValidateParamsArray(connectedShape, connectionSite);
            InvokerService.InvokeInternal.ExecuteMethod(this, "EndConnect", paramsArray, modifiers);
            if (paramsArray[0] is MarshalByRefObject)
                connectedShape = Factory.CreateKnownObjectFromComProxy<NetOffice.WordApi.Shape>(this, paramsArray[0], typeof(NetOffice.WordApi.Shape));
            else
                connectedShape = null;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void EndDisconnect()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "EndDisconnect");
		}

		#endregion

		#pragma warning restore
	}
}
