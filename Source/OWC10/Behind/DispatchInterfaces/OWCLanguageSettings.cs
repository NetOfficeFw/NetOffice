using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.OWC10Api;

namespace NetOffice.OWC10Api.Behind
{
	/// <summary>
	/// DispatchInterface OWCLanguageSettings 
	/// SupportByVersion OWC10, 1
	/// </summary>
	[SupportByVersion("OWC10", 1)]
	[EntityType(EntityType.IsDispatchInterface)]
 	public class OWCLanguageSettings : COMObject, NetOffice.OWC10Api.OWCLanguageSettings
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
                    _contractType = typeof(NetOffice.OWC10Api.OWCLanguageSettings);
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
                    _type = typeof(OWCLanguageSettings);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public OWCLanguageSettings() : base()
		{

		}

		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("OWC10", 1), ProxyResult]
		public virtual object Application
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "Application");
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		/// <param name="id">NetOffice.OWC10Api.Enums.MsoAppLanguageID id</param>
		[SupportByVersion("OWC10", 1)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual Int32 get_LanguageID(NetOffice.OWC10Api.Enums.MsoAppLanguageID id)
		{
			return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "LanguageID", id);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Alias for get_LanguageID
		/// </summary>
		/// <param name="id">NetOffice.OWC10Api.Enums.MsoAppLanguageID id</param>
		[SupportByVersion("OWC10", 1), Redirect("get_LanguageID")]
		public virtual Int32 LanguageID(NetOffice.OWC10Api.Enums.MsoAppLanguageID id)
		{
			return get_LanguageID(id);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		/// <param name="lid">NetOffice.OWC10Api.Enums.MsoLanguageID lid</param>
		[SupportByVersion("OWC10", 1)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual bool get_LanguagePreferredForEditing(NetOffice.OWC10Api.Enums.MsoLanguageID lid)
		{
			return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "LanguagePreferredForEditing", lid);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Alias for get_LanguagePreferredForEditing
		/// </summary>
		/// <param name="lid">NetOffice.OWC10Api.Enums.MsoLanguageID lid</param>
		[SupportByVersion("OWC10", 1), Redirect("get_LanguagePreferredForEditing")]
		public virtual bool LanguagePreferredForEditing(NetOffice.OWC10Api.Enums.MsoLanguageID lid)
		{
			return get_LanguagePreferredForEditing(lid);
		}

		#endregion

		#region Methods

		#endregion

		#pragma warning restore
	}
}

