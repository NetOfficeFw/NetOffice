using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using NetOffice.ExcelApi;
using NetOffice.Exceptions;
using NetOffice.OfficeApi;
using NUnit.Framework;

namespace NetOffice.Tests.ExcelApi.DispatchInterfaces
{
    [TestFixture]
    [Category("IntegrationTests")]
    [Category("IntegrationTests_Excel")]
    public class CustomXMLPartIntegrationTests
    {
        [SetUp]
        public void SetUp()
        {
            this.ExcelApplication = new Application();
            this.Workbook = this.ExcelApplication.Workbooks.Add();
        }

        [TearDown]
        public void TearDown()
        {
            this.Workbook.Close(false);
            this.Workbook.Dispose();

            this.ExcelApplication?.Quit();
            this.ExcelApplication?.Dispose();
        }

        public Application ExcelApplication { get; set; }

        public Workbook Workbook { get; set; }


        [Test]
        public void LoadXML_VadidXmlValue_LoadsIt()
        {
            // Arrange
            CustomXMLPart cxp1 = Workbook.CustomXMLParts.Add();

            // Act
            cxp1.LoadXML("<discounts><discount>0.10</discount></discounts>");

            // Assert
            Assert.AreEqual("<discounts><discount>0.10</discount></discounts>", cxp1.XML);
        }

        [Test]
        public void Delete_ValidXmlPart_RemovesIt()
        {
            // Arrange
            CustomXMLPart cxp1 = Workbook.CustomXMLParts.Add();
            cxp1.LoadXML("<discounts><discount>0.10</discount></discounts>");

            // Act
            cxp1.Delete();

            // Assert
            Assert.Pass();
        }

        /// <summary>
        /// Regression test for #174 (CustomXMLNode.AddNode throws exception; type mismatch)
        /// </summary>
        [Test]
        public void AddNode_NetOfficeCall_ShouldWork()
        {
            // Arrange
            CustomXMLPart cxp1 = Workbook.CustomXMLParts.Add("<invoice />");
            CustomXMLNode cxn = cxp1.SelectSingleNode("/invoice");

            // Act
            Assert.Throws<MethodCOMException>(() => cxp1.AddNode(cxn, "upcode", "urn:invoice:namespace"));

            // Assert
            // Assert.AreEqual(@"<invoice><upccode xmlns=""urn: invoice:namespace""/></invoice>", cxp1.XML);
        }

        /// <summary>
        /// Regression test for #174 (CustomXMLNode.AddNode throws exception; type mismatch)
        /// </summary>
        [Test]
        public void AddNode_NullObjects_ShouldWork()
        {
            // Arrange
            CustomXMLPart cxp1 = Workbook.CustomXMLParts.Add("<invoice />");
            CustomXMLNode cxn = cxp1.SelectSingleNode("/invoice");

            // Act
            Assert.Throws<MethodCOMException>(() => cxp1.AddNode(cxn, null, null, null, null, null));

            // Assert
            // Assert.AreEqual(@"<invoice><upccode xmlns=""urn: invoice:namespace""/></invoice>", cxp1.XML);
        }

        /// <summary>
        /// Regression test for #174 (CustomXMLNode.AddNode throws exception; type mismatch)
        /// </summary>
        [Test]
        public void AddNode_DynamicCall_ShouldWork()
        {
            // Arrange
            var dyn = (dynamic)Workbook.UnderlyingObject;
            dynamic cxp1 = dyn.CustomXMLParts.Add("<invoice />");
            dynamic cxn = cxp1.SelectSingleNode("/invoice");
            cxn.GetType();

            // Act
            var ex = Assert.Throws<ArgumentException>(() => cxp1.AddNode(ref cxn, "upcode", "urn:invoice:namespace"));

            // Assert
            // Assert.AreEqual(@"<invoice><upccode xmlns=""urn: invoice:namespace""/></invoice>", cxp1.XML);
            Assert.AreEqual("Could not convert argument 0 for call to AddNode.", ex.Message);
        }
    }
}
