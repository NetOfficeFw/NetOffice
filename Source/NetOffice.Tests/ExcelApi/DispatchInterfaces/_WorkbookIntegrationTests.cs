using System;
using NetOffice.ExcelApi;
using NetOffice.Exceptions;
using NUnit.Framework;

namespace NetOffice.Tests.ExcelApi.DispatchInterfaces
{
    [TestFixture]
    [Category("IntegrationTests")]
    [Category("IntegrationTests_Excel")]
    public class _WorkbookIntegrationTests
    {
        [SetUp]
        public void SetUp()
        {
            this.ExcelApplication = new Application();
        }

        [TearDown]
        public void TearDown()
        {
            this.ExcelApplication?.Quit();
            this.ExcelApplication?.Dispose();
        }

        public Application ExcelApplication { get; set; }

        [Test(Description = "When a new workbook is created, the default value for the AutoSaveOn property is False and the property is disabled.")]
        public void AutoSaveOn_NewWorkbookDocument_ReturnsFalse()
        {
            // Arrange
            using (var workbook = this.ExcelApplication.Workbooks.Add())
            {
                // Act
                var actual = workbook.AutoSaveOn;

                // Assert
                Assert.IsFalse(actual, "Workbook.AutoSaveOn property should be false.");
            }
        }

        [Test(Description = "When workbook is not hosted in cloud, the AutoSaveOn property is disabled and setting it will result in an error.")]
        public void AutoSaveOn_SetValueForLocalWorkbook_Fails([Values]bool autoSaveOnValue)
        {
            // Arrange
            using (var workbook = this.ExcelApplication.Workbooks.Add())
            {
                // Act && Assert
                var ex = Assert.Throws<PropertySetCOMException>(() => workbook.AutoSaveOn = autoSaveOnValue);

                Assert.AreEqual("Failed to proceed PropertySet on Excel.Workbook=>AutoSaveOn.", ex.Message);
            }
        }
    }
}
