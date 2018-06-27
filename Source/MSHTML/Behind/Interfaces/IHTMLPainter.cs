using System.Reflection;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.MSHTMLApi;

namespace NetOffice.MSHTMLApi.Behind
{
	/// <summary>
	/// Interface IHTMLPainter 
	/// SupportByVersion MSHTML, 4
	/// </summary>
	[SupportByVersion("MSHTML", 4)]
	[EntityType(EntityType.IsInterface)]
 	public class IHTMLPainter : COMObject, NetOffice.MSHTMLApi.IHTMLPainter
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
                    _contractType = typeof(NetOffice.MSHTMLApi.IHTMLPainter);
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
                    _type = typeof(IHTMLPainter);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public IHTMLPainter() : base()
		{

		}

		#endregion
		
		#region Properties

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="rcBounds">tagRECT rcBounds</param>
		/// <param name="rcUpdate">tagRECT rcUpdate</param>
		/// <param name="lDrawFlags">Int32 lDrawFlags</param>
		/// <param name="hdc">_RemotableHandle hdc</param>
		/// <param name="pvDrawObject">object pvDrawObject</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 Draw(tagRECT rcBounds, tagRECT rcUpdate, Int32 lDrawFlags, _RemotableHandle hdc, object pvDrawObject)
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "Draw", new object[]{ rcBounds, rcUpdate, lDrawFlags, hdc, pvDrawObject });
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="size">tagSIZE size</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 onresize(tagSIZE size)
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "onresize", size);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pInfo">_HTML_PAINTER_INFO pInfo</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 GetPainterInfo(out _HTML_PAINTER_INFO pInfo)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(true);
			pInfo = new NetOffice.MSHTMLApi._HTML_PAINTER_INFO();
			object[] paramsArray = Invoker.ValidateParamsArray(pInfo);
			object returnItem = Invoker.MethodReturn(this, "GetPainterInfo", paramsArray, modifiers);
			pInfo = (_HTML_PAINTER_INFO)paramsArray[0];
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pt">tagPOINT pt</param>
		/// <param name="pbHit">Int32 pbHit</param>
		/// <param name="plPartID">Int32 plPartID</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 HitTestPoint(tagPOINT pt, out Int32 pbHit, out Int32 plPartID)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false,true,true);
			pbHit = 0;
			plPartID = 0;
			object[] paramsArray = Invoker.ValidateParamsArray(pt, pbHit, plPartID);
			object returnItem = Invoker.MethodReturn(this, "HitTestPoint", paramsArray, modifiers);
			pbHit = (Int32)paramsArray[1];
			plPartID = (Int32)paramsArray[2];
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		#endregion

		#pragma warning restore
	}
}

