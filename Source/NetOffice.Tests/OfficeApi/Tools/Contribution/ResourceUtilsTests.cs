using System;
using System.Drawing.Imaging;
using System.Reflection;
using NUnit.Framework;
using NetOffice.OfficeApi.Tools.Contribution;
using NetOffice.Tests.Helpers;

namespace NetOffice.Tests.OfficeApi.Tools.Contribution
{
    [TestFixture]
    public class ResourceUtilsTests
    {
        [Test]
        public void ReadString_ResourceName_ReturnsTheImageFromCurrentAssembly()
        {
            // Arrange
            var fakeComObject = new TestableComObjectStub();
            var commons = new CommonUtils(fakeComObject, Assembly.GetExecutingAssembly());
            var utils = commons.Resource;

            // Act
            var actualText = utils.ReadString("NetOffice.Tests.Data.SampleText.txt");

            // Assert
            Assert.AreEqual("Sample text resource file.", actualText);
        }

        [Test]
        public void ReadString_ResourceNameAndAssembly_ReturnsTheImageFromResource()
        {
            // Arrange
            var fakeComObject = new TestableComObjectStub();
            var commons = new CommonUtils(fakeComObject);
            var utils = commons.Resource;

            // Act
            var actualText = utils.ReadString("NetOffice.Tests.Data.SampleText.txt", Assembly.GetExecutingAssembly());

            // Assert
            Assert.AreEqual("Sample text resource file.", actualText);
        }

        [Test]
        public void ReadImage_ResourceName_ReturnsTheImageFromCurrentAssembly()
        {
            // Arrange
            var fakeComObject = new TestableComObjectStub();
            var commons = new CommonUtils(fakeComObject, Assembly.GetExecutingAssembly());
            var utils = commons.Resource;

            // Act
            var actualImage = utils.ReadImage("NetOffice.Tests.Data.SampleImage.png");

            // Assert
            Assert.IsNotNull(actualImage);
            Assert.AreEqual(4, actualImage.Height);
            Assert.AreEqual(4, actualImage.Width);
            Assert.AreEqual(ImageFormat.Png, actualImage.RawFormat);
        }

        [Test]
        public void ReadImage_ResourceNameAndAssembly_ReturnsTheImageFromResource()
        {
            // Arrange
            var fakeComObject = new TestableComObjectStub();
            var commons = new CommonUtils(fakeComObject);
            var utils = commons.Resource;

            // Act
            var actualImage = utils.ReadImage("NetOffice.Tests.Data.SampleImage.png", Assembly.GetExecutingAssembly());

            // Assert
            Assert.IsNotNull(actualImage);
            Assert.AreEqual(4, actualImage.Height);
            Assert.AreEqual(4, actualImage.Width);
            Assert.AreEqual(ImageFormat.Png, actualImage.RawFormat);
        }

        [Test]
        public void ReadIcon_ResourceName_ReturnsTheIconFromCurrentAssembly()
        {
            // Arrange
            var fakeComObject = new TestableComObjectStub();
            var commons = new CommonUtils(fakeComObject, Assembly.GetExecutingAssembly());
            var utils = commons.Resource;

            // Act
            var actualIcon = utils.ReadIcon("NetOffice.Tests.Data.SampleIcon.ico");

            // Assert
            Assert.IsNotNull(actualIcon);
            Assert.AreEqual(16, actualIcon.Height);
            Assert.AreEqual(16, actualIcon.Width);
        }

        [Test]
        public void ReadIcon_ResourceNameAndAssembly_ReturnsTheIconFromResource()
        {
            // Arrange
            var fakeComObject = new TestableComObjectStub();
            var commons = new CommonUtils(fakeComObject);
            var utils = commons.Resource;

            // Act
            var actualIcon = utils.ReadIcon("NetOffice.Tests.Data.SampleIcon.ico", Assembly.GetExecutingAssembly());

            // Assert
            Assert.IsNotNull(actualIcon);
            Assert.AreEqual(16, actualIcon.Height);
            Assert.AreEqual(16, actualIcon.Width);
        }
    }
}
