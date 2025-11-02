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

            // Cleanup
            presentation.Close();
        }

        [Test]
        public void AutoSaveOn_PresentationSavedInOneDrive_FeatureAutoSaveIsOn_ReturnsTrue()
        {
            // Arrange
            var presentation = PowerPointApp.Presentations.Open(this.Docs.AutoSaveOn_PresentationSavedInOneDrive_FeatureAutoSaveIsOn, false);

            // Act
            var actualValue = presentation.AutoSaveOn;

            // Assert
            Assert.IsTrue(actualValue);

            // Cleanup
            presentation.Close();
        }

        public class Documents
        {
            public string AutoSaveOn_PresentationSavedLocally_FeatureAutoSaveIsOff {  get; set; }
            public string AutoSaveOn_PresentationSavedInOneDrive_FeatureAutoSaveIsOn { get; set; }

            public Documents(TestContext context)
            {
                this.AutoSaveOn_PresentationSavedLocally_FeatureAutoSaveIsOff = Path.Combine(context.TestDirectory, @"PowerPointApi\Docs", "AutoSaveOn_PresentationSavedLocally_FeatureAutoSaveIsOff.pptx");

                // Public link: https://1drv.ms/p/c/8cd14a64b99957bc/EVX1bUNmBtxDvfvg-aNQjKkBX35RJ3BJ6_2ey2Ox5gnIfA?e=DVyExX
                this.AutoSaveOn_PresentationSavedInOneDrive_FeatureAutoSaveIsOn = @"https://d.docs.live.net/8CD14A64B99957BC/Developer/NetOfficeFw/PowerPoint/AutoSaveOn_PresentationSavedInOneDrive_FeatureAutoSaveIsOn.pptx";
            }
        }
    }
}
