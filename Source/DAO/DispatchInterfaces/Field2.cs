using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.DAOApi
{
	/// <summary>
	/// DispatchInterface Field2 
	/// SupportByVersion DAO, 12.0
	/// </summary>
	[SupportByVersion("DAO", 12.0)]
	[EntityType(EntityType.IsDispatchInterface)]
	[TypeId("00000054-0000-0010-8000-00AA006D2EA4")]
	public interface Field2 : _Field
	{
		#region Properties

		/// <summary>
		/// SupportByVersion DAO 12.0
		/// Get
		/// </summary>
		[SupportByVersion("DAO", 12.0)]
		new NetOffice.DAOApi.Properties Properties { get; }

		/// <summary>
		/// SupportByVersion DAO 12.0
		/// Get
		/// </summary>
		[SupportByVersion("DAO", 12.0)]
		NetOffice.DAOApi.ComplexType ComplexType { get; }

		/// <summary>
		/// SupportByVersion DAO 12.0
		/// Get
		/// </summary>
		[SupportByVersion("DAO", 12.0)]
		bool IsComplex { get; }

		/// <summary>
		/// SupportByVersion DAO 12.0
		/// Get/Set
		/// </summary>
		[SupportByVersion("DAO", 12.0)]
		bool AppendOnly { get; set; }

		/// <summary>
		/// SupportByVersion DAO 12.0
		/// Get/Set
		/// </summary>
		[SupportByVersion("DAO", 12.0)]
		string Expression { get; set; }

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion DAO 12.0
		/// </summary>
		/// <param name="fileName">string fileName</param>
		[SupportByVersion("DAO", 12.0)]
		void LoadFromFile(string fileName);

		/// <summary>
		/// SupportByVersion DAO 12.0
		/// </summary>
		/// <param name="fileName">string fileName</param>
		[SupportByVersion("DAO", 12.0)]
		void SaveToFile(string fileName);

		#endregion
	}
}
