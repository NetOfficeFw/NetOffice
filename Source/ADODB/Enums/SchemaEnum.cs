using System;
using NetOffice;
using NetOffice.Attributes;
namespace NetOffice.ADODBApi.Enums
{
	 /// <summary>
	 /// SupportByVersion ADODB 2.1, 2.5
	 /// </summary>
	[SupportByVersion("ADODB", 2.1,2.5)]
	[EntityType(EntityType.IsEnum)]
	public enum SchemaEnum
	{
		 /// <summary>
		 /// SupportByVersion ADODB 2.1, 2.5
		 /// </summary>
		 /// <remarks>-1</remarks>
		 [SupportByVersion("ADODB", 2.1,2.5)]
		 adSchemaProviderSpecific = -1,

		 /// <summary>
		 /// SupportByVersion ADODB 2.1, 2.5
		 /// </summary>
		 /// <remarks>0</remarks>
		 [SupportByVersion("ADODB", 2.1,2.5)]
		 adSchemaAsserts = 0,

		 /// <summary>
		 /// SupportByVersion ADODB 2.1, 2.5
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersion("ADODB", 2.1,2.5)]
		 adSchemaCatalogs = 1,

		 /// <summary>
		 /// SupportByVersion ADODB 2.1, 2.5
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersion("ADODB", 2.1,2.5)]
		 adSchemaCharacterSets = 2,

		 /// <summary>
		 /// SupportByVersion ADODB 2.1, 2.5
		 /// </summary>
		 /// <remarks>3</remarks>
		 [SupportByVersion("ADODB", 2.1,2.5)]
		 adSchemaCollations = 3,

		 /// <summary>
		 /// SupportByVersion ADODB 2.1, 2.5
		 /// </summary>
		 /// <remarks>4</remarks>
		 [SupportByVersion("ADODB", 2.1,2.5)]
		 adSchemaColumns = 4,

		 /// <summary>
		 /// SupportByVersion ADODB 2.1, 2.5
		 /// </summary>
		 /// <remarks>5</remarks>
		 [SupportByVersion("ADODB", 2.1,2.5)]
		 adSchemaCheckConstraints = 5,

		 /// <summary>
		 /// SupportByVersion ADODB 2.1, 2.5
		 /// </summary>
		 /// <remarks>6</remarks>
		 [SupportByVersion("ADODB", 2.1,2.5)]
		 adSchemaConstraintColumnUsage = 6,

		 /// <summary>
		 /// SupportByVersion ADODB 2.1, 2.5
		 /// </summary>
		 /// <remarks>7</remarks>
		 [SupportByVersion("ADODB", 2.1,2.5)]
		 adSchemaConstraintTableUsage = 7,

		 /// <summary>
		 /// SupportByVersion ADODB 2.1, 2.5
		 /// </summary>
		 /// <remarks>8</remarks>
		 [SupportByVersion("ADODB", 2.1,2.5)]
		 adSchemaKeyColumnUsage = 8,

		 /// <summary>
		 /// SupportByVersion ADODB 2.1, 2.5
		 /// </summary>
		 /// <remarks>9</remarks>
		 [SupportByVersion("ADODB", 2.1,2.5)]
		 adSchemaReferentialContraints = 9,

		 /// <summary>
		 /// SupportByVersion ADODB 2.1, 2.5
		 /// </summary>
		 /// <remarks>9</remarks>
		 [SupportByVersion("ADODB", 2.1,2.5)]
		 adSchemaReferentialConstraints = 9,

		 /// <summary>
		 /// SupportByVersion ADODB 2.1, 2.5
		 /// </summary>
		 /// <remarks>10</remarks>
		 [SupportByVersion("ADODB", 2.1,2.5)]
		 adSchemaTableConstraints = 10,

		 /// <summary>
		 /// SupportByVersion ADODB 2.1, 2.5
		 /// </summary>
		 /// <remarks>11</remarks>
		 [SupportByVersion("ADODB", 2.1,2.5)]
		 adSchemaColumnsDomainUsage = 11,

		 /// <summary>
		 /// SupportByVersion ADODB 2.1, 2.5
		 /// </summary>
		 /// <remarks>12</remarks>
		 [SupportByVersion("ADODB", 2.1,2.5)]
		 adSchemaIndexes = 12,

		 /// <summary>
		 /// SupportByVersion ADODB 2.1, 2.5
		 /// </summary>
		 /// <remarks>13</remarks>
		 [SupportByVersion("ADODB", 2.1,2.5)]
		 adSchemaColumnPrivileges = 13,

		 /// <summary>
		 /// SupportByVersion ADODB 2.1, 2.5
		 /// </summary>
		 /// <remarks>14</remarks>
		 [SupportByVersion("ADODB", 2.1,2.5)]
		 adSchemaTablePrivileges = 14,

		 /// <summary>
		 /// SupportByVersion ADODB 2.1, 2.5
		 /// </summary>
		 /// <remarks>15</remarks>
		 [SupportByVersion("ADODB", 2.1,2.5)]
		 adSchemaUsagePrivileges = 15,

		 /// <summary>
		 /// SupportByVersion ADODB 2.1, 2.5
		 /// </summary>
		 /// <remarks>16</remarks>
		 [SupportByVersion("ADODB", 2.1,2.5)]
		 adSchemaProcedures = 16,

		 /// <summary>
		 /// SupportByVersion ADODB 2.1, 2.5
		 /// </summary>
		 /// <remarks>17</remarks>
		 [SupportByVersion("ADODB", 2.1,2.5)]
		 adSchemaSchemata = 17,

		 /// <summary>
		 /// SupportByVersion ADODB 2.1, 2.5
		 /// </summary>
		 /// <remarks>18</remarks>
		 [SupportByVersion("ADODB", 2.1,2.5)]
		 adSchemaSQLLanguages = 18,

		 /// <summary>
		 /// SupportByVersion ADODB 2.1, 2.5
		 /// </summary>
		 /// <remarks>19</remarks>
		 [SupportByVersion("ADODB", 2.1,2.5)]
		 adSchemaStatistics = 19,

		 /// <summary>
		 /// SupportByVersion ADODB 2.1, 2.5
		 /// </summary>
		 /// <remarks>20</remarks>
		 [SupportByVersion("ADODB", 2.1,2.5)]
		 adSchemaTables = 20,

		 /// <summary>
		 /// SupportByVersion ADODB 2.1, 2.5
		 /// </summary>
		 /// <remarks>21</remarks>
		 [SupportByVersion("ADODB", 2.1,2.5)]
		 adSchemaTranslations = 21,

		 /// <summary>
		 /// SupportByVersion ADODB 2.1, 2.5
		 /// </summary>
		 /// <remarks>22</remarks>
		 [SupportByVersion("ADODB", 2.1,2.5)]
		 adSchemaProviderTypes = 22,

		 /// <summary>
		 /// SupportByVersion ADODB 2.1, 2.5
		 /// </summary>
		 /// <remarks>23</remarks>
		 [SupportByVersion("ADODB", 2.1,2.5)]
		 adSchemaViews = 23,

		 /// <summary>
		 /// SupportByVersion ADODB 2.1, 2.5
		 /// </summary>
		 /// <remarks>24</remarks>
		 [SupportByVersion("ADODB", 2.1,2.5)]
		 adSchemaViewColumnUsage = 24,

		 /// <summary>
		 /// SupportByVersion ADODB 2.1, 2.5
		 /// </summary>
		 /// <remarks>25</remarks>
		 [SupportByVersion("ADODB", 2.1,2.5)]
		 adSchemaViewTableUsage = 25,

		 /// <summary>
		 /// SupportByVersion ADODB 2.1, 2.5
		 /// </summary>
		 /// <remarks>26</remarks>
		 [SupportByVersion("ADODB", 2.1,2.5)]
		 adSchemaProcedureParameters = 26,

		 /// <summary>
		 /// SupportByVersion ADODB 2.1, 2.5
		 /// </summary>
		 /// <remarks>27</remarks>
		 [SupportByVersion("ADODB", 2.1,2.5)]
		 adSchemaForeignKeys = 27,

		 /// <summary>
		 /// SupportByVersion ADODB 2.1, 2.5
		 /// </summary>
		 /// <remarks>28</remarks>
		 [SupportByVersion("ADODB", 2.1,2.5)]
		 adSchemaPrimaryKeys = 28,

		 /// <summary>
		 /// SupportByVersion ADODB 2.1, 2.5
		 /// </summary>
		 /// <remarks>29</remarks>
		 [SupportByVersion("ADODB", 2.1,2.5)]
		 adSchemaProcedureColumns = 29,

		 /// <summary>
		 /// SupportByVersion ADODB 2.1, 2.5
		 /// </summary>
		 /// <remarks>30</remarks>
		 [SupportByVersion("ADODB", 2.1,2.5)]
		 adSchemaDBInfoKeywords = 30,

		 /// <summary>
		 /// SupportByVersion ADODB 2.1, 2.5
		 /// </summary>
		 /// <remarks>31</remarks>
		 [SupportByVersion("ADODB", 2.1,2.5)]
		 adSchemaDBInfoLiterals = 31,

		 /// <summary>
		 /// SupportByVersion ADODB 2.1, 2.5
		 /// </summary>
		 /// <remarks>32</remarks>
		 [SupportByVersion("ADODB", 2.1,2.5)]
		 adSchemaCubes = 32,

		 /// <summary>
		 /// SupportByVersion ADODB 2.1, 2.5
		 /// </summary>
		 /// <remarks>33</remarks>
		 [SupportByVersion("ADODB", 2.1,2.5)]
		 adSchemaDimensions = 33,

		 /// <summary>
		 /// SupportByVersion ADODB 2.1, 2.5
		 /// </summary>
		 /// <remarks>34</remarks>
		 [SupportByVersion("ADODB", 2.1,2.5)]
		 adSchemaHierarchies = 34,

		 /// <summary>
		 /// SupportByVersion ADODB 2.1, 2.5
		 /// </summary>
		 /// <remarks>35</remarks>
		 [SupportByVersion("ADODB", 2.1,2.5)]
		 adSchemaLevels = 35,

		 /// <summary>
		 /// SupportByVersion ADODB 2.1, 2.5
		 /// </summary>
		 /// <remarks>36</remarks>
		 [SupportByVersion("ADODB", 2.1,2.5)]
		 adSchemaMeasures = 36,

		 /// <summary>
		 /// SupportByVersion ADODB 2.1, 2.5
		 /// </summary>
		 /// <remarks>37</remarks>
		 [SupportByVersion("ADODB", 2.1,2.5)]
		 adSchemaProperties = 37,

		 /// <summary>
		 /// SupportByVersion ADODB 2.1, 2.5
		 /// </summary>
		 /// <remarks>38</remarks>
		 [SupportByVersion("ADODB", 2.1,2.5)]
		 adSchemaMembers = 38,

		 /// <summary>
		 /// SupportByVersion ADODB 2.1, 2.5
		 /// </summary>
		 /// <remarks>39</remarks>
		 [SupportByVersion("ADODB", 2.1,2.5)]
		 adSchemaTrustees = 39
	}
}