using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.VisioApi;

namespace NetOffice.VisioApi.Behind
{
	/// <summary>
	/// DispatchInterface IVisEventProc 
	/// SupportByVersion Visio, 11,12,14,15,16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/ff769310(v=office.14).aspx </remarks>
	[SupportByVersion("Visio", 11,12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface)]
 	public class IVisEventProc : COMObject, NetOffice.VisioApi.IVisEventProc
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
                    _contractType = typeof(NetOffice.VisioApi.IVisEventProc);
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
                    _type = typeof(IVisEventProc);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public IVisEventProc() : base()
		{

		}

		#endregion
		
		#region Properties

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/ff768483(v=office.14).aspx </remarks>
		/// <param name="nEventCode">Int16 nEventCode</param>
		/// <param name="pSourceObj">object pSourceObj</param>
		/// <param name="nEventID">Int32 nEventID</param>
		/// <param name="nEventSeqNum">Int32 nEventSeqNum</param>
		/// <param name="pSubjectObj">object pSubjectObj</param>
		/// <param name="vMoreInfo">object vMoreInfo</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual object VisEventProc(Int16 nEventCode, object pSourceObj, Int32 nEventID, Int32 nEventSeqNum, object pSubjectObj, object vMoreInfo)
		{
			return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "VisEventProc", new object[]{ nEventCode, pSourceObj, nEventID, nEventSeqNum, pSubjectObj, vMoreInfo });
		}

		#endregion

		#pragma warning restore
	}
}

