using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.AccessApi
{
	/// <summary>
	/// DispatchInterface _SqmProxyReceiver 
	/// SupportByVersion Access, 15, 16
	/// </summary>
	[SupportByVersion("Access", 15, 16)]
	[EntityType(EntityType.IsDispatchInterface)]
 	public class _SqmProxyReceiver : COMObject
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
                    _type = typeof(_SqmProxyReceiver);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public _SqmProxyReceiver(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public _SqmProxyReceiver(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _SqmProxyReceiver(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _SqmProxyReceiver(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _SqmProxyReceiver(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _SqmProxyReceiver(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _SqmProxyReceiver() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _SqmProxyReceiver(string progId) : base(progId)
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
		public void SetDataPoint(UIntPtr id, UIntPtr dwValue)
		{
			 Factory.ExecuteMethod(this, "SetDataPoint", id, dwValue);
		}

		/// <summary>
		/// SupportByVersion Access 15,16
		/// </summary>
		/// <param name="id">UIntPtr id</param>
		/// <param name="dwValue">UIntPtr dwValue</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 15, 16)]
		public void SetDataPointMax(UIntPtr id, UIntPtr dwValue)
		{
			 Factory.ExecuteMethod(this, "SetDataPointMax", id, dwValue);
		}

		/// <summary>
		/// SupportByVersion Access 15,16
		/// </summary>
		/// <param name="id">UIntPtr id</param>
		/// <param name="dwValue">UIntPtr dwValue</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 15, 16)]
		public void SetDataPointMin(UIntPtr id, UIntPtr dwValue)
		{
			 Factory.ExecuteMethod(this, "SetDataPointMin", id, dwValue);
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
		public void CreateStream(UIntPtr id, UIntPtr type, UIntPtr width, UIntPtr maxRows)
		{
			 Factory.ExecuteMethod(this, "CreateStream", id, type, width, maxRows);
		}

		/// <summary>
		/// SupportByVersion Access 15,16
		/// </summary>
		/// <param name="id">UIntPtr id</param>
		/// <param name="dw1">UIntPtr dw1</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 15, 16)]
		public void AddStreamData1(UIntPtr id, UIntPtr dw1)
		{
			 Factory.ExecuteMethod(this, "AddStreamData1", id, dw1);
		}

		/// <summary>
		/// SupportByVersion Access 15,16
		/// </summary>
		/// <param name="id">UIntPtr id</param>
		/// <param name="dw1">UIntPtr dw1</param>
		/// <param name="dw2">UIntPtr dw2</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 15, 16)]
		public void AddStreamData2(UIntPtr id, UIntPtr dw1, UIntPtr dw2)
		{
			 Factory.ExecuteMethod(this, "AddStreamData2", id, dw1, dw2);
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
		public void AddStreamData8(UIntPtr id, UIntPtr dw1, UIntPtr dw2, UIntPtr dw3, UIntPtr dw4, UIntPtr dw5, UIntPtr dw6, UIntPtr dw7, UIntPtr dw8)
		{
			 Factory.ExecuteMethod(this, "AddStreamData8", new object[]{ id, dw1, dw2, dw3, dw4, dw5, dw6, dw7, dw8 });
		}

		#endregion

		#pragma warning restore
	}
}
