using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.OutlookApi;

namespace NetOffice.OutlookApi.Behind
{
	/// <summary>
	/// DispatchInterface _FormRegionStartup 
	/// SupportByVersion Outlook, 12,14,15,16
	/// </summary>
	[SupportByVersion("Outlook", 12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface), BaseType]
 	public class _FormRegionStartup : COMObject, NetOffice.OutlookApi._FormRegionStartup
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
                    _contractType = typeof(NetOffice.OutlookApi._FormRegionStartup);
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
                    _type = typeof(_FormRegionStartup);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public _FormRegionStartup() : base()
		{

		}

		#endregion
		
		#region Properties

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff866042.aspx </remarks>
		/// <param name="formRegionName">string formRegionName</param>
		/// <param name="item">object item</param>
		/// <param name="lCID">Int32 lCID</param>
		/// <param name="formRegionMode">NetOffice.OutlookApi.Enums.OlFormRegionMode formRegionMode</param>
		/// <param name="formRegionSize">NetOffice.OutlookApi.Enums.OlFormRegionSize formRegionSize</param>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public virtual object GetFormRegionStorage(string formRegionName, object item, Int32 lCID, NetOffice.OutlookApi.Enums.OlFormRegionMode formRegionMode, NetOffice.OutlookApi.Enums.OlFormRegionSize formRegionSize)
		{
			return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "GetFormRegionStorage", new object[]{ formRegionName, item, lCID, formRegionMode, formRegionSize });
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff869072.aspx </remarks>
		/// <param name="formRegion">NetOffice.OutlookApi.FormRegion formRegion</param>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public virtual void BeforeFormRegionShow(NetOffice.OutlookApi.FormRegion formRegion)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "BeforeFormRegionShow", formRegion);
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff869502.aspx </remarks>
		/// <param name="formRegionName">string formRegionName</param>
		/// <param name="lCID">Int32 lCID</param>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public virtual object GetFormRegionManifest(string formRegionName, Int32 lCID)
		{
			return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "GetFormRegionManifest", formRegionName, lCID);
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff868914.aspx </remarks>
		/// <param name="formRegionName">string formRegionName</param>
		/// <param name="lCID">Int32 lCID</param>
		/// <param name="icon">NetOffice.OutlookApi.Enums.OlFormRegionIcon icon</param>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public virtual object GetFormRegionIcon(string formRegionName, Int32 lCID, NetOffice.OutlookApi.Enums.OlFormRegionIcon icon)
		{
			return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "GetFormRegionIcon", formRegionName, lCID, icon);
		}

		#endregion

		#pragma warning restore
	}
}

