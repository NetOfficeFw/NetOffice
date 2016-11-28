using System;
using NetOffice;
namespace NetOffice.PublisherApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Publisher 14, 15, 16
	 /// </summary>
	[SupportByVersionAttribute("Publisher", 14,15,16)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum PbWizardPageType
	{
		 /// <summary>
		 /// SupportByVersion Publisher 14, 15, 16
		 /// </summary>
		 /// <remarks>-1</remarks>
		 [SupportByVersionAttribute("Publisher", 14,15,16)]
		 pbWizardPageTypeNone = -1,

		 /// <summary>
		 /// SupportByVersion Publisher 14, 15, 16
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersionAttribute("Publisher", 14,15,16)]
		 pbWizardPageTypeNewsletter3Stories = 1,

		 /// <summary>
		 /// SupportByVersion Publisher 14, 15, 16
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersionAttribute("Publisher", 14,15,16)]
		 pbWizardPageTypeNewsletterCalendar = 2,

		 /// <summary>
		 /// SupportByVersion Publisher 14, 15, 16
		 /// </summary>
		 /// <remarks>15</remarks>
		 [SupportByVersionAttribute("Publisher", 14,15,16)]
		 pbWizardPageTypeNewsletterOrderForm = 15,

		 /// <summary>
		 /// SupportByVersion Publisher 14, 15, 16
		 /// </summary>
		 /// <remarks>16</remarks>
		 [SupportByVersionAttribute("Publisher", 14,15,16)]
		 pbWizardPageTypeNewsletterResponseForm = 16,

		 /// <summary>
		 /// SupportByVersion Publisher 14, 15, 16
		 /// </summary>
		 /// <remarks>17</remarks>
		 [SupportByVersionAttribute("Publisher", 14,15,16)]
		 pbWizardPageTypeNewsletterSignupForm = 17,

		 /// <summary>
		 /// SupportByVersion Publisher 14, 15, 16
		 /// </summary>
		 /// <remarks>18</remarks>
		 [SupportByVersionAttribute("Publisher", 14,15,16)]
		 pbWizardPageTypeCatalogOneColumnText = 18,

		 /// <summary>
		 /// SupportByVersion Publisher 14, 15, 16
		 /// </summary>
		 /// <remarks>19</remarks>
		 [SupportByVersionAttribute("Publisher", 14,15,16)]
		 pbWizardPageTypeCatalogOneColumnTextPicture = 19,

		 /// <summary>
		 /// SupportByVersion Publisher 14, 15, 16
		 /// </summary>
		 /// <remarks>20</remarks>
		 [SupportByVersionAttribute("Publisher", 14,15,16)]
		 pbWizardPageTypeCatalogTwoColumnsText = 20,

		 /// <summary>
		 /// SupportByVersion Publisher 14, 15, 16
		 /// </summary>
		 /// <remarks>21</remarks>
		 [SupportByVersionAttribute("Publisher", 14,15,16)]
		 pbWizardPageTypeCatalogTwoColumnsTextPicture = 21,

		 /// <summary>
		 /// SupportByVersion Publisher 14, 15, 16
		 /// </summary>
		 /// <remarks>22</remarks>
		 [SupportByVersionAttribute("Publisher", 14,15,16)]
		 pbWizardPageTypeCatalogCalendar = 22,

		 /// <summary>
		 /// SupportByVersion Publisher 14, 15, 16
		 /// </summary>
		 /// <remarks>23</remarks>
		 [SupportByVersionAttribute("Publisher", 14,15,16)]
		 pbWizardPageTypeCatalogTableOfContents = 23,

		 /// <summary>
		 /// SupportByVersion Publisher 14, 15, 16
		 /// </summary>
		 /// <remarks>24</remarks>
		 [SupportByVersionAttribute("Publisher", 14,15,16)]
		 pbWizardPageTypeCatalogFeaturedItem = 24,

		 /// <summary>
		 /// SupportByVersion Publisher 14, 15, 16
		 /// </summary>
		 /// <remarks>25</remarks>
		 [SupportByVersionAttribute("Publisher", 14,15,16)]
		 pbWizardPageTypeCatalogTwoItemsAlignedPictures = 25,

		 /// <summary>
		 /// SupportByVersion Publisher 14, 15, 16
		 /// </summary>
		 /// <remarks>26</remarks>
		 [SupportByVersionAttribute("Publisher", 14,15,16)]
		 pbWizardPageTypeCatalogTwoItemsOffsetPictures = 26,

		 /// <summary>
		 /// SupportByVersion Publisher 14, 15, 16
		 /// </summary>
		 /// <remarks>27</remarks>
		 [SupportByVersionAttribute("Publisher", 14,15,16)]
		 pbWizardPageTypeCatalogThreeItemsAlignedPictures = 27,

		 /// <summary>
		 /// SupportByVersion Publisher 14, 15, 16
		 /// </summary>
		 /// <remarks>28</remarks>
		 [SupportByVersionAttribute("Publisher", 14,15,16)]
		 pbWizardPageTypeCatalogThreeItemsOffsetPictures = 28,

		 /// <summary>
		 /// SupportByVersion Publisher 14, 15, 16
		 /// </summary>
		 /// <remarks>29</remarks>
		 [SupportByVersionAttribute("Publisher", 14,15,16)]
		 pbWizardPageTypeCatalogThreeItemsStackedPictures = 29,

		 /// <summary>
		 /// SupportByVersion Publisher 14, 15, 16
		 /// </summary>
		 /// <remarks>30</remarks>
		 [SupportByVersionAttribute("Publisher", 14,15,16)]
		 pbWizardPageTypeCatalogFourItemsAlignedPictures = 30,

		 /// <summary>
		 /// SupportByVersion Publisher 14, 15, 16
		 /// </summary>
		 /// <remarks>31</remarks>
		 [SupportByVersionAttribute("Publisher", 14,15,16)]
		 pbWizardPageTypeCatalogFourItemsOffsetPictures = 31,

		 /// <summary>
		 /// SupportByVersion Publisher 14, 15, 16
		 /// </summary>
		 /// <remarks>32</remarks>
		 [SupportByVersionAttribute("Publisher", 14,15,16)]
		 pbWizardPageTypeCatalogFourItemsSquaredPictures = 32,

		 /// <summary>
		 /// SupportByVersion Publisher 14, 15, 16
		 /// </summary>
		 /// <remarks>33</remarks>
		 [SupportByVersionAttribute("Publisher", 14,15,16)]
		 pbWizardPageTypeCatalogEightItemsOneColumn = 33,

		 /// <summary>
		 /// SupportByVersion Publisher 14, 15, 16
		 /// </summary>
		 /// <remarks>34</remarks>
		 [SupportByVersionAttribute("Publisher", 14,15,16)]
		 pbWizardPageTypeCatalogEightItemsTwoColumns = 34,

		 /// <summary>
		 /// SupportByVersion Publisher 14, 15, 16
		 /// </summary>
		 /// <remarks>35</remarks>
		 [SupportByVersionAttribute("Publisher", 14,15,16)]
		 pbWizardPageTypeCatalogBlank = 35,

		 /// <summary>
		 /// SupportByVersion Publisher 14, 15, 16
		 /// </summary>
		 /// <remarks>36</remarks>
		 [SupportByVersionAttribute("Publisher", 14,15,16)]
		 pbWizardPageTypeCatalogForm = 36,

		 /// <summary>
		 /// SupportByVersion Publisher 14, 15, 16
		 /// </summary>
		 /// <remarks>501</remarks>
		 [SupportByVersionAttribute("Publisher", 14,15,16)]
		 pbWizardPageTypeWebAboutUs = 501,

		 /// <summary>
		 /// SupportByVersion Publisher 14, 15, 16
		 /// </summary>
		 /// <remarks>502</remarks>
		 [SupportByVersionAttribute("Publisher", 14,15,16)]
		 pbWizardPageTypeWebInformational = 502,

		 /// <summary>
		 /// SupportByVersion Publisher 14, 15, 16
		 /// </summary>
		 /// <remarks>503</remarks>
		 [SupportByVersionAttribute("Publisher", 14,15,16)]
		 pbWizardPageTypeWebList = 503,

		 /// <summary>
		 /// SupportByVersion Publisher 14, 15, 16
		 /// </summary>
		 /// <remarks>504</remarks>
		 [SupportByVersionAttribute("Publisher", 14,15,16)]
		 pbWizardPageTypeWebCalendarPage = 504,

		 /// <summary>
		 /// SupportByVersion Publisher 14, 15, 16
		 /// </summary>
		 /// <remarks>505</remarks>
		 [SupportByVersionAttribute("Publisher", 14,15,16)]
		 pbWizardPageTypeWebContactUs = 505,

		 /// <summary>
		 /// SupportByVersion Publisher 14, 15, 16
		 /// </summary>
		 /// <remarks>506</remarks>
		 [SupportByVersionAttribute("Publisher", 14,15,16)]
		 pbWizardPageTypeWebEmployeeList = 506,

		 /// <summary>
		 /// SupportByVersion Publisher 14, 15, 16
		 /// </summary>
		 /// <remarks>507</remarks>
		 [SupportByVersionAttribute("Publisher", 14,15,16)]
		 pbWizardPageTypeWebEmployee = 507,

		 /// <summary>
		 /// SupportByVersion Publisher 14, 15, 16
		 /// </summary>
		 /// <remarks>508</remarks>
		 [SupportByVersionAttribute("Publisher", 14,15,16)]
		 pbWizardPageTypeWebFAQ = 508,

		 /// <summary>
		 /// SupportByVersion Publisher 14, 15, 16
		 /// </summary>
		 /// <remarks>509</remarks>
		 [SupportByVersionAttribute("Publisher", 14,15,16)]
		 pbWizardPageTypeWebHome = 509,

		 /// <summary>
		 /// SupportByVersion Publisher 14, 15, 16
		 /// </summary>
		 /// <remarks>510</remarks>
		 [SupportByVersionAttribute("Publisher", 14,15,16)]
		 pbWizardPageTypeWebJobs = 510,

		 /// <summary>
		 /// SupportByVersion Publisher 14, 15, 16
		 /// </summary>
		 /// <remarks>511</remarks>
		 [SupportByVersionAttribute("Publisher", 14,15,16)]
		 pbWizardPageTypeWebLegal = 511,

		 /// <summary>
		 /// SupportByVersion Publisher 14, 15, 16
		 /// </summary>
		 /// <remarks>512</remarks>
		 [SupportByVersionAttribute("Publisher", 14,15,16)]
		 pbWizardPageTypeWebArticle = 512,

		 /// <summary>
		 /// SupportByVersion Publisher 14, 15, 16
		 /// </summary>
		 /// <remarks>513</remarks>
		 [SupportByVersionAttribute("Publisher", 14,15,16)]
		 pbWizardPageTypeWebPhoto = 513,

		 /// <summary>
		 /// SupportByVersion Publisher 14, 15, 16
		 /// </summary>
		 /// <remarks>514</remarks>
		 [SupportByVersionAttribute("Publisher", 14,15,16)]
		 pbWizardPageTypeWebPhotoGallery = 514,

		 /// <summary>
		 /// SupportByVersion Publisher 14, 15, 16
		 /// </summary>
		 /// <remarks>515</remarks>
		 [SupportByVersionAttribute("Publisher", 14,15,16)]
		 pbWizardPageTypeWebProduct = 515,

		 /// <summary>
		 /// SupportByVersion Publisher 14, 15, 16
		 /// </summary>
		 /// <remarks>516</remarks>
		 [SupportByVersionAttribute("Publisher", 14,15,16)]
		 pbWizardPageTypeWebProductList = 516,

		 /// <summary>
		 /// SupportByVersion Publisher 14, 15, 16
		 /// </summary>
		 /// <remarks>517</remarks>
		 [SupportByVersionAttribute("Publisher", 14,15,16)]
		 pbWizardPageTypeWebProjectList = 517,

		 /// <summary>
		 /// SupportByVersion Publisher 14, 15, 16
		 /// </summary>
		 /// <remarks>518</remarks>
		 [SupportByVersionAttribute("Publisher", 14,15,16)]
		 pbWizardPageTypeWebLinks = 518,

		 /// <summary>
		 /// SupportByVersion Publisher 14, 15, 16
		 /// </summary>
		 /// <remarks>519</remarks>
		 [SupportByVersionAttribute("Publisher", 14,15,16)]
		 pbWizardPageTypeWebSeminar = 519,

		 /// <summary>
		 /// SupportByVersion Publisher 14, 15, 16
		 /// </summary>
		 /// <remarks>520</remarks>
		 [SupportByVersionAttribute("Publisher", 14,15,16)]
		 pbWizardPageTypeWebServiceList = 520,

		 /// <summary>
		 /// SupportByVersion Publisher 14, 15, 16
		 /// </summary>
		 /// <remarks>521</remarks>
		 [SupportByVersionAttribute("Publisher", 14,15,16)]
		 pbWizardPageTypeWebService = 521,

		 /// <summary>
		 /// SupportByVersion Publisher 14, 15, 16
		 /// </summary>
		 /// <remarks>522</remarks>
		 [SupportByVersionAttribute("Publisher", 14,15,16)]
		 pbWizardPageTypeWebSpecial = 522,

		 /// <summary>
		 /// SupportByVersion Publisher 14, 15, 16
		 /// </summary>
		 /// <remarks>524</remarks>
		 [SupportByVersionAttribute("Publisher", 14,15,16)]
		 pbWizardPageTypeWebBlank = 524,

		 /// <summary>
		 /// SupportByVersion Publisher 14, 15, 16
		 /// </summary>
		 /// <remarks>525</remarks>
		 [SupportByVersionAttribute("Publisher", 14,15,16)]
		 pbWizardPageTypeWebOrderForm = 525,

		 /// <summary>
		 /// SupportByVersion Publisher 14, 15, 16
		 /// </summary>
		 /// <remarks>526</remarks>
		 [SupportByVersionAttribute("Publisher", 14,15,16)]
		 pbWizardPageTypeWebResponseForm = 526,

		 /// <summary>
		 /// SupportByVersion Publisher 14, 15, 16
		 /// </summary>
		 /// <remarks>527</remarks>
		 [SupportByVersionAttribute("Publisher", 14,15,16)]
		 pbWizardPageTypeWebSignupForm = 527,

		 /// <summary>
		 /// SupportByVersion Publisher 14, 15, 16
		 /// </summary>
		 /// <remarks>800</remarks>
		 [SupportByVersionAttribute("Publisher", 14,15,16)]
		 pbWizardPageTypeWebCalendarWithLinks = 800,

		 /// <summary>
		 /// SupportByVersion Publisher 14, 15, 16
		 /// </summary>
		 /// <remarks>801</remarks>
		 [SupportByVersionAttribute("Publisher", 14,15,16)]
		 pbWizardPageTypeWebProductsWithLinks = 801,

		 /// <summary>
		 /// SupportByVersion Publisher 14, 15, 16
		 /// </summary>
		 /// <remarks>802</remarks>
		 [SupportByVersionAttribute("Publisher", 14,15,16)]
		 pbWizardPageTypeWebEmployeesWithLinks = 802,

		 /// <summary>
		 /// SupportByVersion Publisher 14, 15, 16
		 /// </summary>
		 /// <remarks>803</remarks>
		 [SupportByVersionAttribute("Publisher", 14,15,16)]
		 pbWizardPageTypeWebServicesWithLinks = 803,

		 /// <summary>
		 /// SupportByVersion Publisher 14, 15, 16
		 /// </summary>
		 /// <remarks>804</remarks>
		 [SupportByVersionAttribute("Publisher", 14,15,16)]
		 pbWizardPageTypeWebProjectsWithLinks = 804,

		 /// <summary>
		 /// SupportByVersion Publisher 14, 15, 16
		 /// </summary>
		 /// <remarks>805</remarks>
		 [SupportByVersionAttribute("Publisher", 14,15,16)]
		 pbWizardPageTypeWebPhotosWithLinks = 805
	}
}