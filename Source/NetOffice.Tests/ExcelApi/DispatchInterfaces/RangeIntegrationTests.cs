using System;
using System.Collections.Generic;
using NetOffice.ExcelApi;
using NUnit.Framework;
using Application = NetOffice.ExcelApi.Application;

namespace NetOffice.Tests.ExcelApi.DispatchInterfaces
{
    [TestFixture]
    [Category("IntegrationTests")]
    [Category("IntegrationTests_Excel")]
    public class RangeIntegrationTests
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


        /// <summary>
        /// Regression test for #177 (Range.Offset[] does not work)
        /// </summary>
        [Test]
        public void Offset_Overload1_NegativeIndexInNativeObject_ReturnsExpectedRange()
        {
            // Arrange
            using (var workbook = this.ExcelApplication.Workbooks.Add())
            {
                var sheet = workbook.Worksheets[1] as Worksheet;

                // Act
                var range = sheet.Cells[5, 5];
                var cell1 = (dynamic)range.UnderlyingObject;
                var cell2 = cell1.Offset[-1];

                // Assert
                Assert.AreEqual("$E$5", cell1.Address);
                Assert.AreEqual("$E$4", cell2.Address);
            }
        }

        /// <summary>
        /// Regression test for #177 (Range.Offset[] does not work)
        /// </summary>
        [Test]
        public void Offset_Overload2_NegativeIndexInNativeObject_ReturnsExpectedRange()
        {
            // Arrange
            using (var workbook = this.ExcelApplication.Workbooks.Add())
            {
                var sheet = workbook.Worksheets[1] as Worksheet;

                // Act
                var range = sheet.Cells[5, 5];
                var cell1 = (dynamic)range.UnderlyingObject;
                var cell2 = cell1.Offset[-1, 0];

                // Assert
                Assert.AreEqual("$E$5", cell1.Address);
                Assert.AreEqual("$E$4", cell2.Address);
            }
        }

        /// <summary>
        /// Regression test for #177 (Range.Offset[] does not work)
        /// </summary>
        [Test]
        public void Offset_Indexer0_NegativeIndexInNetOfficeObject_ReturnsExpectedRange()
        {
            // Arrange
            using (var workbook = this.ExcelApplication.Workbooks.Add())
            {
                var sheet = workbook.Worksheets[1] as Worksheet;

                // Act
                var cell1 = sheet.Cells[5, 5];
                var cell2 = cell1.Offset[0, 0];

                // Assert
                Assert.AreEqual("$E$5", cell1.Address);
                Assert.AreEqual("$D$4", cell2.Address);
            }
        }

        /// <summary>
        /// Regression test for #177 (Range.Offset[] does not work)
        /// </summary>
        [Test]
        public void Offset_Indexer1_NegativeIndexInNetOfficeObject_ReturnsExpectedRange()
        {
            // Arrange
            using (var workbook = this.ExcelApplication.Workbooks.Add())
            {
                var sheet = workbook.Worksheets[1] as Worksheet;

                // Act
                var cell1 = sheet.Cells[5, 5];
                var cell2 = cell1.Offset[-1];

                // Assert
                Assert.AreEqual("$E$5", cell1.Address);
                Assert.AreEqual("$E$3", cell2.Address);
            }
        }

        /// <summary>
        /// Regression test for #177 (Range.Offset[] does not work)
        /// </summary>
        [Test]
        public void Offset_Method1_NegativeIndexInNetOfficeObject_ReturnsExpectedRange()
        {
            // Arrange
            using (var workbook = this.ExcelApplication.Workbooks.Add())
            {
                var sheet = workbook.Worksheets[1] as Worksheet;

                // Act
                var cell1 = sheet.Cells[5, 5];
                var cell2 = cell1.Offset(-1);

                // Assert
                Assert.AreEqual("$E$5", cell1.Address);
                Assert.AreEqual("$E$4", cell2.Address);
            }
        }

        /// <summary>
        /// Regression test for #177 (Range.Offset[] does not work)
        /// </summary>
        [Test]
        public void Offset_Method2_NegativeIndexInNetOfficeObject_ReturnsExpectedRange()
        {
            // Arrange
            using (var workbook = this.ExcelApplication.Workbooks.Add())
            {
                var sheet = workbook.Worksheets[1] as Worksheet;

                // Act
                var cell1 = sheet.Cells[5, 5];
                var cell2 = cell1.Offset(-1, 0);

                // Assert
                Assert.AreEqual("$E$5", cell1.Address);
                Assert.AreEqual("$E$4", cell2.Address);
            }
        }
    }
}
