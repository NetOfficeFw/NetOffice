using System.Reflection;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.VisioApi
{
	/// <summary>
	/// Interface LPVISIOCURVE 
	/// SupportByVersion Visio, 11,12,14,15,16
	/// </summary>
	[SupportByVersion("Visio", 11,12,14,15,16)]
	[EntityType(EntityType.IsInterface)]
 	public class LPVISIOCURVE : COMObject
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
                    _type = typeof(LPVISIOCURVE);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public LPVISIOCURVE(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public LPVISIOCURVE(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public LPVISIOCURVE(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public LPVISIOCURVE(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public LPVISIOCURVE(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public LPVISIOCURVE(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public LPVISIOCURVE() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public LPVISIOCURVE(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[BaseResult]
		public NetOffice.VisioApi.IVApplication Application
		{
			get
			{
				return Factory.ExecuteBaseReferencePropertyGet<NetOffice.VisioApi.IVApplication>(this, "Application");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public Int16 ObjectType
		{
			get
			{
				return Factory.ExecuteInt16PropertyGet(this, "ObjectType");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public Int16 Closed
		{
			get
			{
				return Factory.ExecuteInt16PropertyGet(this, "Closed");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public Double Start
		{
			get
			{
				return Factory.ExecuteDoublePropertyGet(this, "Start");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public Double End
		{
			get
			{
				return Factory.ExecuteDoublePropertyGet(this, "End");
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="tolerance">Double tolerance</param>
		/// <param name="xyArray">Double[] xyArray</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public void Points(Double tolerance, out Double[] xyArray)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false,true);
			xyArray = null;
			object[] paramsArray = Invoker.ValidateParamsArray(tolerance, (object)xyArray);
			Invoker.Method(this, "Points", paramsArray, modifiers);
			xyArray = (Double[])paramsArray[1];
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="t">Double t</param>
		/// <param name="x">Double x</param>
		/// <param name="y">Double y</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public void Point(Double t, out Double x, out Double y)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false,true,true);
			x = 0;
			y = 0;
			object[] paramsArray = Invoker.ValidateParamsArray(t, x, y);
			Invoker.Method(this, "Point", paramsArray, modifiers);
			x = (Double)paramsArray[1];
			y = (Double)paramsArray[2];
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="t">Double t</param>
		/// <param name="n">Int16 n</param>
		/// <param name="x">Double x</param>
		/// <param name="y">Double y</param>
		/// <param name="dxdt">Double dxdt</param>
		/// <param name="dydt">Double dydt</param>
		/// <param name="ddxdt">Double ddxdt</param>
		/// <param name="ddydt">Double ddydt</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public void PointAndDerivatives(Double t, Int16 n, out Double x, out Double y, out Double dxdt, out Double dydt, out Double ddxdt, out Double ddydt)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false,false,true,true,true,true,true,true);
			x = 0;
			y = 0;
			dxdt = 0;
			dydt = 0;
			ddxdt = 0;
			ddydt = 0;
			object[] paramsArray = Invoker.ValidateParamsArray(t, n, x, y, dxdt, dydt, ddxdt, ddydt);
			Invoker.Method(this, "PointAndDerivatives", paramsArray, modifiers);
			x = (Double)paramsArray[2];
			y = (Double)paramsArray[3];
			dxdt = (Double)paramsArray[4];
			dydt = (Double)paramsArray[5];
			ddxdt = (Double)paramsArray[6];
			ddydt = (Double)paramsArray[7];
		}

		#endregion

		#pragma warning restore
	}
}
