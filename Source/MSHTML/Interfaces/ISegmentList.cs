using System.Reflection;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.MSHTMLApi
{
	/// <summary>
	/// Interface ISegmentList 
	/// SupportByVersion MSHTML, 4
	/// </summary>
	[SupportByVersion("MSHTML", 4)]
	[EntityType(EntityType.IsInterface)]
 	public class ISegmentList : COMObject
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
                    _type = typeof(ISegmentList);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public ISegmentList(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public ISegmentList(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ISegmentList(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ISegmentList(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ISegmentList(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ISegmentList(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ISegmentList() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ISegmentList(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="ppIIter">NetOffice.MSHTMLApi.ISegmentListIterator ppIIter</param>
		[SupportByVersion("MSHTML", 4)]
		public Int32 CreateIterator(out NetOffice.MSHTMLApi.ISegmentListIterator ppIIter)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(true);
			ppIIter = null;
			object[] paramsArray = Invoker.ValidateParamsArray(ppIIter);
			object returnItem = Invoker.MethodReturn(this, "CreateIterator", paramsArray, modifiers);
            if (paramsArray[0] is MarshalByRefObject)
                ppIIter = new NetOffice.MSHTMLApi.ISegmentListIterator(this, paramsArray[0]);
            else
                ppIIter = null;
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="peType">NetOffice.MSHTMLApi.Enums._SELECTION_TYPE peType</param>
		[SupportByVersion("MSHTML", 4)]
		public Int32 GetType(out NetOffice.MSHTMLApi.Enums._SELECTION_TYPE peType)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(true);
			peType = 0;
			object[] paramsArray = Invoker.ValidateParamsArray(peType);
			object returnItem = Invoker.MethodReturn(this, "GetType", paramsArray, modifiers);
			peType = (NetOffice.MSHTMLApi.Enums._SELECTION_TYPE)paramsArray[0];
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pfEmpty">Int32 pfEmpty</param>
		[SupportByVersion("MSHTML", 4)]
		public Int32 IsEmpty(out Int32 pfEmpty)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(true);
			pfEmpty = 0;
			object[] paramsArray = Invoker.ValidateParamsArray(pfEmpty);
			object returnItem = Invoker.MethodReturn(this, "IsEmpty", paramsArray, modifiers);
			pfEmpty = (Int32)paramsArray[0];
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		#endregion

		#pragma warning restore
	}
}
