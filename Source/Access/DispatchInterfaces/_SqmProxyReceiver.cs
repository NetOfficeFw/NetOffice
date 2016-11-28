using System;
using NetRuntimeSystem = System;
using System.Runtime.InteropServices;
using System.Runtime.CompilerServices;
using System.ComponentModel;
using System.Reflection;
using System.Collections.Generic;
using NetOffice;
namespace NetOffice.AccessApi
{
	///<summary>
	/// DispatchInterface _SqmProxyReceiver 
	/// SupportByVersion Access, 15, 16
	///</summary>
	[SupportByVersionAttribute("Access", 15, 16)]
	[EntityTypeAttribute(EntityType.IsDispatchInterface)]
	public class _SqmProxyReceiver : COMObject
	{
		#pragma warning disable
		#region Type Information

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
        
		#region Construction

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
		
		/// <param name="progId">registered ProgID</param>
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
		/// 
		/// </summary>
		/// <param name="id">UIntPtr id</param>
		/// <param name="dwValue">UIntPtr dwValue</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Access", 15, 16)]
		public void SetDataPoint(UIntPtr id, UIntPtr dwValue)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(id, dwValue);
			Invoker.Method(this, "SetDataPoint", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 15,16
		/// 
		/// </summary>
		/// <param name="id">UIntPtr id</param>
		/// <param name="dwValue">UIntPtr dwValue</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Access", 15, 16)]
		public void SetDataPointMax(UIntPtr id, UIntPtr dwValue)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(id, dwValue);
			Invoker.Method(this, "SetDataPointMax", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 15,16
		/// 
		/// </summary>
		/// <param name="id">UIntPtr id</param>
		/// <param name="dwValue">UIntPtr dwValue</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Access", 15, 16)]
		public void SetDataPointMin(UIntPtr id, UIntPtr dwValue)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(id, dwValue);
			Invoker.Method(this, "SetDataPointMin", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 15,16
		/// 
		/// </summary>
		/// <param name="id">UIntPtr id</param>
		/// <param name="type">UIntPtr Type</param>
		/// <param name="width">UIntPtr Width</param>
		/// <param name="maxRows">UIntPtr maxRows</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Access", 15, 16)]
		public void CreateStream(UIntPtr id, UIntPtr type, UIntPtr width, UIntPtr maxRows)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(id, type, width, maxRows);
			Invoker.Method(this, "CreateStream", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 15,16
		/// 
		/// </summary>
		/// <param name="id">UIntPtr id</param>
		/// <param name="dw1">UIntPtr dw1</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Access", 15, 16)]
		public void AddStreamData1(UIntPtr id, UIntPtr dw1)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(id, dw1);
			Invoker.Method(this, "AddStreamData1", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 15,16
		/// 
		/// </summary>
		/// <param name="id">UIntPtr id</param>
		/// <param name="dw1">UIntPtr dw1</param>
		/// <param name="dw2">UIntPtr dw2</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Access", 15, 16)]
		public void AddStreamData2(UIntPtr id, UIntPtr dw1, UIntPtr dw2)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(id, dw1, dw2);
			Invoker.Method(this, "AddStreamData2", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 15,16
		/// 
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
		[SupportByVersionAttribute("Access", 15, 16)]
		public void AddStreamData8(UIntPtr id, UIntPtr dw1, UIntPtr dw2, UIntPtr dw3, UIntPtr dw4, UIntPtr dw5, UIntPtr dw6, UIntPtr dw7, UIntPtr dw8)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(id, dw1, dw2, dw3, dw4, dw5, dw6, dw7, dw8);
			Invoker.Method(this, "AddStreamData8", paramsArray);
		}

		#endregion
		#pragma warning restore
	}
}