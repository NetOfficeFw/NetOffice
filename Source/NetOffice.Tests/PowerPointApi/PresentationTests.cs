using System;
using NUnit.Framework;
using NetOffice.PowerPointApi;
using System.IO;

namespace NetOffice.Tests.PowerPointApi
{
    [TestFixture]
    [Category("IntegrationTests")]
    [Category("IntegrationTests_PowerPoint")]
    public class PresentationTests
    {
        public Application PowerPointApp { get; set; }
        public Documents Docs { get; set; }

        [SetUp]
        public void SetUp()
        {
            this.PowerPointApp = new Application();
            this.Docs = new Documents(TestContext.CurrentContext);
        }

        [TearDown]
        public void TearDown()
        {
            this.PowerPointApp?.Quit();
            this.PowerPointApp?.Dispose();
        }

        [Test]
        public void AutoSaveOn_PresentationSavedLocally_FeatureAutoSaveIsOff_ReturnsFalse()
        {
            // Arrange
            var presentation = PowerPointApp.Presentations.Open(this.Docs.AutoSaveOn_PresentationSavedLocally_FeatureAutoSaveIsOff, false);

            // Act
            var actualValue = presentation.AutoSaveOn;

            // Assert
            Assert.IsFalse(actualValue);
        }

        public class Documents
        {
            public string AutoSaveOn_PresentationSavedLocally_FeatureAutoSaveIsOff {  get; set; }

            public Documents(TestContext context)
            {
                this.AutoSaveOn_PresentationSavedLocally_FeatureAutoSaveIsOff = Path.Combine(context.TestDirectory, @"PowerPointApi\Docs", "AutoSaveOn_PresentationSavedLocally_FeatureAutoSaveIsOff.pptx");
            }
        }
    }
}
