using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.AccessApi;

namespace NetOffice.AccessApi.Behind
{
	/// <summary>
	/// DispatchInterface _SqmProxyReceiver 
	/// SupportByVersion Access, 15, 16
	/// </summary>
	[SupportByVersion("Access", 15, 16)]
	[EntityType(EntityType.IsDispatchInterface)]
 	public class _SqmProxyReceiver : COMObject, NetOffice.AccessApi._SqmProxyReceiver
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
                    _contractType = typeof(NetOffice.AccessApi._SqmProxyReceiver);
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
                    _type = typeof(_SqmProxyReceiver);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public _SqmProxyReceiver() : base()
		{

		}

		#endregion
		
		#region Properties

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Access 15,16
		/// </summary>
		/// <param name="id">UIntPtr id</param>
		/// <param name="dwValue">UIntPtr dwValue</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 15, 16)]
		public virtual void SetDataPoint(UIntPtr id, UIntPtr dwValue)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SetDataPoint", id, dwValue);
		}

		/// <summary>
		/// SupportByVersion Access 15,16
		/// </summary>
		/// <param name="id">UIntPtr id</param>
		/// <param name="dwValue">UIntPtr dwValue</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 15, 16)]
		public virtual void SetDataPointMax(UIntPtr id, UIntPtr dwValue)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SetDataPointMax", id, dwValue);
		}

		/// <summary>
		/// SupportByVersion Access 15,16
		/// </summary>
		/// <param name="id">UIntPtr id</param>
		/// <param name="dwValue">UIntPtr dwValue</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 15, 16)]
		public virtual void SetDataPointMin(UIntPtr id, UIntPtr dwValue)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SetDataPointMin", id, dwValue);
		}

		/// <summary>
		/// SupportByVersion Access 15,16
		/// </summary>
		/// <param name="id">UIntPtr id</param>
		/// <param name="type">UIntPtr type</param>
		/// <param name="width">UIntPtr width</param>
		/// <param name="maxRows">UIntPtr maxRows</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 15, 16)]
		public virtual void CreateStream(UIntPtr id, UIntPtr type, UIntPtr width, UIntPtr maxRows)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "CreateStream", id, type, width, maxRows);
		}

		/// <summary>
		/// SupportByVersion Access 15,16
		/// </summary>
		/// <param name="id">UIntPtr id</param>
		/// <param name="dw1">UIntPtr dw1</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 15, 16)]
		public virtual void AddStreamData1(UIntPtr id, UIntPtr dw1)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "AddStreamData1", id, dw1);
		}

		/// <summary>
		/// SupportByVersion Access 15,16
		/// </summary>
		/// <param name="id">UIntPtr id</param>
		/// <param name="dw1">UIntPtr dw1</param>
		/// <param name="dw2">UIntPtr dw2</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 15, 16)]
		public virtual void AddStreamData2(UIntPtr id, UIntPtr dw1, UIntPtr dw2)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "AddStreamData2", id, dw1, dw2);
		}

		/// <summary>
		/// SupportByVersion Access 15,16
		/// </summary>
		/// <param name="id">UIntPtr id</param>
		/// <param name="dw1">UIntPtr dw1</param>
		/// <param name="dw2">UIntPtr dw2</param>
		/// <param name="dw3">UIntPtr dw3</param>
		/// <param name="dw4">UIntPtr dw4</param>
		/// <param name="dw5">UIntPtr dw5</param>
		/// <param name="dw6">UIntPtr dw6</param>
		/// <param name="dw7">UIntPtr dw7</param>
		/// <param name="dw8">UIntPtr dw8</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 15, 16)]
		public virtual void AddStreamData8(UIntPtr id, UIntPtr dw1, UIntPtr dw2, UIntPtr dw3, UIntPtr dw4, UIntPtr dw5, UIntPtr dw6, UIntPtr dw7, UIntPtr dw8)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "AddStreamData8", new object[]{ id, dw1, dw2, dw3, dw4, dw5, dw6, dw7, dw8 });
		}

		#endregion

		#pragma warning restore
	}
}

