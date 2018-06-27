using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.OWC10Api;

namespace NetOffice.OWC10Api.Behind
{
	/// <summary>
	/// DispatchInterface DataPage 
	/// SupportByVersion OWC10, 1
	/// </summary>
	[SupportByVersion("OWC10", 1)]
	[EntityType(EntityType.IsDispatchInterface)]
 	public class DataPage : COMObject, NetOffice.OWC10Api.DataPage
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
                    _contractType = typeof(NetOffice.OWC10Api.DataPage);
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
                    _type = typeof(DataPage);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public DataPage() : base()
		{

		}

		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public virtual NetOffice.OWC10Api.GroupLevel GroupLevel
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OWC10Api.GroupLevel>(this, "GroupLevel", typeof(NetOffice.OWC10Api.GroupLevel));
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public virtual NetOffice.ADODBApi.Recordset Recordset
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.ADODBApi.Recordset>(this, "Recordset", typeof(NetOffice.ADODBApi.Recordset));
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public virtual object Filter
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Filter");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "Filter", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public virtual bool IsFilterOn
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "IsFilterOn");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "IsFilterOn", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public virtual NetOffice.OWC10Api.Section FirstSection
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OWC10Api.Section>(this, "FirstSection", typeof(NetOffice.OWC10Api.Section));
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public virtual void MoveFirst()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "MoveFirst");
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public virtual void MoveLast()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "MoveLast");
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public virtual void MoveNext()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "MoveNext");
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public virtual void MovePrevious()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "MovePrevious");
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public virtual void NewRecord()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "NewRecord");
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public virtual void DeleteRecord()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "DeleteRecord");
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public virtual void Save()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Save");
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public virtual void Undo()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Undo");
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public virtual void Requery()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Requery");
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public virtual void SortAscending()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SortAscending");
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public virtual void SortDescending()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SortDescending");
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public virtual void ApplyFilter()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "ApplyFilter");
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public virtual void ToggleFilter()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "ToggleFilter");
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("OWC10", 1)]
		public virtual void Help()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Help");
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="button">NetOffice.OWC10Api.Enums.NavButtonEnum button</param>
		[SupportByVersion("OWC10", 1)]
		public virtual bool IsButtonEnabled(NetOffice.OWC10Api.Enums.NavButtonEnum button)
		{
			return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "IsButtonEnabled", button);
		}

		#endregion

		#pragma warning restore
	}
}


