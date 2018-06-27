using System.Reflection;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.VisioApi;

namespace NetOffice.VisioApi.Behind
{
	/// <summary>
	/// Interface LPVISIOCURVE 
	/// SupportByVersion Visio, 11,12,14,15,16
	/// </summary>
	[SupportByVersion("Visio", 11,12,14,15,16)]
	[EntityType(EntityType.IsInterface)]
 	public class LPVISIOCURVE : COMObject, NetOffice.VisioApi.LPVISIOCURVE
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
                    _contractType = typeof(NetOffice.VisioApi.LPVISIOCURVE);
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
                    _type = typeof(LPVISIOCURVE);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public LPVISIOCURVE() : base()
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
		public virtual NetOffice.VisioApi.IVApplication Application
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBaseReferencePropertyGet<NetOffice.VisioApi.IVApplication>(this, "Application");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual Int16 ObjectType
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt16PropertyGet(this, "ObjectType");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual Int16 Closed
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt16PropertyGet(this, "Closed");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual Double Start
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteDoublePropertyGet(this, "Start");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual Double End
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteDoublePropertyGet(this, "End");
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
		public virtual void Points(Double tolerance, out Double[] xyArray)
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
		public virtual void Point(Double t, out Double x, out Double y)
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
		public virtual void PointAndDerivatives(Double t, Int16 n, out Double x, out Double y, out Double dxdt, out Double dydt, out Double ddxdt, out Double ddydt)
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

