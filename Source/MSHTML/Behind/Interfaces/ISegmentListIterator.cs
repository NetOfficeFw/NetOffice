using System.Reflection;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.MSHTMLApi;

namespace NetOffice.MSHTMLApi.Behind
{
	/// <summary>
	/// Interface ISegmentListIterator 
	/// SupportByVersion MSHTML, 4
	/// </summary>
	[SupportByVersion("MSHTML", 4)]
	[EntityType(EntityType.IsInterface)]
 	public class ISegmentListIterator : COMObject, NetOffice.MSHTMLApi.ISegmentListIterator
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
                    _contractType = typeof(NetOffice.MSHTMLApi.ISegmentListIterator);
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
                    _type = typeof(ISegmentListIterator);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public ISegmentListIterator() : base()
		{

		}

		#endregion
		
		#region Properties

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="ppISegment">NetOffice.MSHTMLApi.ISegment ppISegment</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 Current(out NetOffice.MSHTMLApi.ISegment ppISegment)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(true);
			ppISegment = null;
			object[] paramsArray = Invoker.ValidateParamsArray(ppISegment);
			object returnItem = Invoker.MethodReturn(this, "Current", paramsArray, modifiers);
            if (paramsArray[0] is MarshalByRefObject)
                ppISegment = Factory.CreateKnownObjectFromComProxy<NetOffice.MSHTMLApi.ISegment>(this, paramsArray[0], typeof(NetOffice.MSHTMLApi.ISegment));
            else
                ppISegment = null;
            return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 First()
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "First");
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 IsDone()
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "IsDone");
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 Advance()
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "Advance");
		}

		#endregion

		#pragma warning restore
	}
}

