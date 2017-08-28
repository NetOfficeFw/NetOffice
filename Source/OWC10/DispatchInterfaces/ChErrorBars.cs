using System.Reflection;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.OWC10Api
{
	/// <summary>
	/// DispatchInterface ChErrorBars 
	/// SupportByVersion OWC10, 1
	/// </summary>
	[SupportByVersion("OWC10", 1)]
	[EntityType(EntityType.IsDispatchInterface)]
 	public class ChErrorBars : COMObject
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
                    _type = typeof(ChErrorBars);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public ChErrorBars(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public ChErrorBars(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ChErrorBars(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ChErrorBars(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ChErrorBars(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ChErrorBars(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ChErrorBars() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ChErrorBars(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public NetOffice.OWC10Api.Enums.ChartEndStyleEnum EndStyle
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OWC10Api.Enums.ChartEndStyleEnum>(this, "EndStyle");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "EndStyle", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public NetOffice.OWC10Api.Enums.ChartErrorBarDirectionEnum Direction
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OWC10Api.Enums.ChartErrorBarDirectionEnum>(this, "Direction");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "Direction", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public Int32 Index
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "Index");
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public NetOffice.OWC10Api.ChLine Line
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OWC10Api.ChLine>(this, "Line", NetOffice.OWC10Api.ChLine.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public NetOffice.OWC10Api.ChSeries Parent
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OWC10Api.ChSeries>(this, "Parent", NetOffice.OWC10Api.ChSeries.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public Double Amount
		{
			get
			{
				return Factory.ExecuteDoublePropertyGet(this, "Amount");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Amount", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public NetOffice.OWC10Api.Enums.ChartErrorBarIncludeEnum Include
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OWC10Api.Enums.ChartErrorBarIncludeEnum>(this, "Include");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "Include", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public NetOffice.OWC10Api.Enums.ChartErrorBarTypeEnum Type
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OWC10Api.Enums.ChartErrorBarTypeEnum>(this, "Type");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "Type", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public Int32 Top
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "Top");
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public Int32 Left
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "Left");
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public Int32 Bottom
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "Bottom");
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public Int32 Right
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "Right");
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public NetOffice.OWC10Api.Enums.ChartSelectionsEnum ObjectType
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OWC10Api.Enums.ChartSelectionsEnum>(this, "ObjectType");
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="dimension">NetOffice.OWC10Api.Enums.ChartErrorBarCustomValuesEnum dimension</param>
		/// <param name="dataSourceIndex">Int32 dataSourceIndex</param>
		/// <param name="dataReference">optional object dataReference</param>
		[SupportByVersion("OWC10", 1)]
		public void SetData(NetOffice.OWC10Api.Enums.ChartErrorBarCustomValuesEnum dimension, Int32 dataSourceIndex, object dataReference)
		{
			 Factory.ExecuteMethod(this, "SetData", dimension, dataSourceIndex, dataReference);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="dimension">NetOffice.OWC10Api.Enums.ChartErrorBarCustomValuesEnum dimension</param>
		/// <param name="dataSourceIndex">Int32 dataSourceIndex</param>
		[CustomMethod]
		[SupportByVersion("OWC10", 1)]
		public void SetData(NetOffice.OWC10Api.Enums.ChartErrorBarCustomValuesEnum dimension, Int32 dataSourceIndex)
		{
			 Factory.ExecuteMethod(this, "SetData", dimension, dataSourceIndex);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="errorBarDirection">NetOffice.OWC10Api.Enums.ChartErrorBarCustomValuesEnum errorBarDirection</param>
		[SupportByVersion("OWC10", 1)]
		public string GetDataReference(NetOffice.OWC10Api.Enums.ChartErrorBarCustomValuesEnum errorBarDirection)
		{
			return Factory.ExecuteStringMethodGet(this, "GetDataReference", errorBarDirection);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="errorBarDirection">NetOffice.OWC10Api.Enums.ChartErrorBarCustomValuesEnum errorBarDirection</param>
		[SupportByVersion("OWC10", 1)]
		public Int32 GetDataSourceIndex(NetOffice.OWC10Api.Enums.ChartErrorBarCustomValuesEnum errorBarDirection)
		{
			return Factory.ExecuteInt32MethodGet(this, "GetDataSourceIndex", errorBarDirection);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="errorBarDirection">NetOffice.OWC10Api.Enums.ChartErrorBarCustomValuesEnum errorBarDirection</param>
		/// <param name="dataSourceIndex">object dataSourceIndex</param>
		/// <param name="dataReference">object dataReference</param>
		[SupportByVersion("OWC10", 1)]
		public void GetData(NetOffice.OWC10Api.Enums.ChartErrorBarCustomValuesEnum errorBarDirection, out object dataSourceIndex, out object dataReference)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false,true,true);
			dataSourceIndex = null;
			dataReference = null;
			object[] paramsArray = Invoker.ValidateParamsArray(errorBarDirection, dataSourceIndex, dataReference);
			Invoker.Method(this, "GetData", paramsArray, modifiers);
			dataSourceIndex = (object)paramsArray[1];
			dataReference = (object)paramsArray[2];
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public void Select()
		{
			 Factory.ExecuteMethod(this, "Select");
		}

		#endregion

		#pragma warning restore
	}
}
