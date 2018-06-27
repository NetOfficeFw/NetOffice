using System.Reflection;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.MSHTMLApi;

namespace NetOffice.MSHTMLApi.Behind
{
	/// <summary>
	/// Interface IElementBehaviorLayout 
	/// SupportByVersion MSHTML, 4
	/// </summary>
	[SupportByVersion("MSHTML", 4)]
	[EntityType(EntityType.IsInterface)]
 	public class IElementBehaviorLayout : COMObject, NetOffice.MSHTMLApi.IElementBehaviorLayout
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
                    _contractType = typeof(NetOffice.MSHTMLApi.IElementBehaviorLayout);
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
                    _type = typeof(IElementBehaviorLayout);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public IElementBehaviorLayout() : base()
		{

		}

		#endregion
		
		#region Properties

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="dwFlags">Int32 dwFlags</param>
		/// <param name="sizeContent">tagSIZE sizeContent</param>
		/// <param name="pptTranslateBy">tagPOINT pptTranslateBy</param>
		/// <param name="pptTopLeft">tagPOINT pptTopLeft</param>
		/// <param name="psizeProposed">tagSIZE psizeProposed</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 GetSize(Int32 dwFlags, tagSIZE sizeContent, tagPOINT pptTranslateBy, tagPOINT pptTopLeft, tagSIZE psizeProposed)
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "GetSize", new object[]{ dwFlags, sizeContent, pptTranslateBy, pptTopLeft, psizeProposed });
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 GetLayoutInfo()
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "GetLayoutInfo");
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="lFlags">Int32 lFlags</param>
		/// <param name="pptTopLeft">tagPOINT pptTopLeft</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 GetPosition(Int32 lFlags, tagPOINT pptTopLeft)
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "GetPosition", lFlags, pptTopLeft);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="psizeIn">tagSIZE psizeIn</param>
		/// <param name="prcOut">tagRECT prcOut</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 MapSize(tagSIZE psizeIn, out tagRECT prcOut)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false,true);
			prcOut = new NetOffice.MSHTMLApi.tagRECT();
			object[] paramsArray = Invoker.ValidateParamsArray(psizeIn, prcOut);
			object returnItem = Invoker.MethodReturn(this, "MapSize", paramsArray, modifiers);
			prcOut = (tagRECT)paramsArray[1];
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		#endregion

		#pragma warning restore
	}
}

