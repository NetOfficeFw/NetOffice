using System.Reflection;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.MSHTMLApi;

namespace NetOffice.MSHTMLApi.Behind
{
	/// <summary>
	/// Interface IMarkupContainer 
	/// SupportByVersion MSHTML, 4
	/// </summary>
	[SupportByVersion("MSHTML", 4)]
	[EntityType(EntityType.IsInterface), BaseType]
 	public class IMarkupContainer : COMObject, NetOffice.MSHTMLApi.IMarkupContainer
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
                    _contractType = typeof(NetOffice.MSHTMLApi.IMarkupContainer);
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
                    _type = typeof(IMarkupContainer);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public IMarkupContainer() : base()
		{

		}

        #endregion

        #region Properties

        #endregion

        #region Methods

        /// <summary>
        /// SupportByVersion MSHTML 4
        /// </summary>
        /// <param name="ppDoc">NetOffice.MSHTMLApi.IHTMLDocument2 ppDoc</param>
        [SupportByVersion("MSHTML", 4)]
        public virtual Int32 OwningDoc(out NetOffice.MSHTMLApi.IHTMLDocument2 ppDoc)
        {
            ParameterModifier[] modifiers = Invoker.CreateParamModifiers(true);
            ppDoc = null;
            object[] paramsArray = Invoker.ValidateParamsArray(ppDoc);
            object returnItem = Invoker.MethodReturn(this, "OwningDoc", paramsArray, modifiers);
            if (paramsArray[0] is MarshalByRefObject)
                ppDoc = Factory.CreateKnownObjectFromComProxy<NetOffice.MSHTMLApi.IHTMLDocument2>(this, paramsArray[0], typeof(NetOffice.MSHTMLApi.IHTMLDocument2));
            else
                ppDoc = null;
            return NetRuntimeSystem.Convert.ToInt32(returnItem);
        }            

		#endregion

		#pragma warning restore
	}
}

