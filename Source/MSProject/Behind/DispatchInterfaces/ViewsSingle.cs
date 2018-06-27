using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.MSProjectApi;

namespace NetOffice.MSProjectApi.Behind
{
	/// <summary>
	/// DispatchInterface ViewsSingle 
	/// SupportByVersion MSProject, 11,12,14
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff920756(v=office.14).aspx </remarks>
	[SupportByVersion("MSProject", 11,12,14)]
	[EntityType(EntityType.IsDispatchInterface)]
 	public class ViewsSingle : Views, NetOffice.MSProjectApi.ViewsSingle
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
                    _contractType = typeof(NetOffice.MSProjectApi.ViewsSingle);
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
                    _type = typeof(ViewsSingle);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public ViewsSingle() : base()
		{

		}

		#endregion
		
		#region Properties

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		/// <param name="name">string name</param>
		/// <param name="screen">optional NetOffice.MSProjectApi.Enums.PjViewScreen Screen = 1</param>
		/// <param name="showInMenu">optional bool ShowInMenu = false</param>
		/// <param name="table">optional object table</param>
		/// <param name="filter">optional object filter</param>
		/// <param name="group">optional object group</param>
		/// <param name="highlightFilt">optional bool HighlightFilt = false</param>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual NetOffice.MSProjectApi.ViewSingle Add(string name, object screen, object showInMenu, object table, object filter, object group, object highlightFilt)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.MSProjectApi.ViewSingle>(this, "Add", typeof(NetOffice.MSProjectApi.ViewSingle), new object[]{ name, screen, showInMenu, table, filter, group, highlightFilt });
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		/// <param name="name">string name</param>
		[CustomMethod]
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual NetOffice.MSProjectApi.ViewSingle Add(string name)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.MSProjectApi.ViewSingle>(this, "Add", typeof(NetOffice.MSProjectApi.ViewSingle), name);
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		/// <param name="name">string name</param>
		/// <param name="screen">optional NetOffice.MSProjectApi.Enums.PjViewScreen Screen = 1</param>
		[CustomMethod]
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual NetOffice.MSProjectApi.ViewSingle Add(string name, object screen)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.MSProjectApi.ViewSingle>(this, "Add", typeof(NetOffice.MSProjectApi.ViewSingle), name, screen);
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		/// <param name="name">string name</param>
		/// <param name="screen">optional NetOffice.MSProjectApi.Enums.PjViewScreen Screen = 1</param>
		/// <param name="showInMenu">optional bool ShowInMenu = false</param>
		[CustomMethod]
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual NetOffice.MSProjectApi.ViewSingle Add(string name, object screen, object showInMenu)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.MSProjectApi.ViewSingle>(this, "Add", typeof(NetOffice.MSProjectApi.ViewSingle), name, screen, showInMenu);
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		/// <param name="name">string name</param>
		/// <param name="screen">optional NetOffice.MSProjectApi.Enums.PjViewScreen Screen = 1</param>
		/// <param name="showInMenu">optional bool ShowInMenu = false</param>
		/// <param name="table">optional object table</param>
		[CustomMethod]
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual NetOffice.MSProjectApi.ViewSingle Add(string name, object screen, object showInMenu, object table)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.MSProjectApi.ViewSingle>(this, "Add", typeof(NetOffice.MSProjectApi.ViewSingle), name, screen, showInMenu, table);
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		/// <param name="name">string name</param>
		/// <param name="screen">optional NetOffice.MSProjectApi.Enums.PjViewScreen Screen = 1</param>
		/// <param name="showInMenu">optional bool ShowInMenu = false</param>
		/// <param name="table">optional object table</param>
		/// <param name="filter">optional object filter</param>
		[CustomMethod]
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual NetOffice.MSProjectApi.ViewSingle Add(string name, object screen, object showInMenu, object table, object filter)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.MSProjectApi.ViewSingle>(this, "Add", typeof(NetOffice.MSProjectApi.ViewSingle), new object[]{ name, screen, showInMenu, table, filter });
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		/// <param name="name">string name</param>
		/// <param name="screen">optional NetOffice.MSProjectApi.Enums.PjViewScreen Screen = 1</param>
		/// <param name="showInMenu">optional bool ShowInMenu = false</param>
		/// <param name="table">optional object table</param>
		/// <param name="filter">optional object filter</param>
		/// <param name="group">optional object group</param>
		[CustomMethod]
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual NetOffice.MSProjectApi.ViewSingle Add(string name, object screen, object showInMenu, object table, object filter, object group)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.MSProjectApi.ViewSingle>(this, "Add", typeof(NetOffice.MSProjectApi.ViewSingle), new object[]{ name, screen, showInMenu, table, filter, group });
		}

		#endregion

		#pragma warning restore
	}
}


