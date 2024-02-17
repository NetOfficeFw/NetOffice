using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Net;
using System.Web;
using System.Drawing;
using System.IO;

namespace Sample.Addin
{
    internal class WebImage
    {
        internal WebImage(string url, Image image)
        {
            Url = url;
            Image = image;
        }

        internal string Url { get; private set; }
        internal Image Image { get; private set; }
    }

    internal class WebImageCollection
    {
        public WebImageCollection()
        {
            Images = new List<WebImage>();
        }
        private List<WebImage> Images { get; set; }

        public Image this[string url]
        {
            get
            {
                foreach (WebImage item in Images)
                {
                    if (item.Url.Equals(url, StringComparison.InvariantCultureIgnoreCase))
                        return item.Image;
                }

                Image newImage = GetImageFromURL(url);
                Images.Add(new WebImage(url, newImage));
                return newImage;
            }
        }

        /// <summary>
        /// Gets the image from URL.
        /// </summary>
        /// <param name="url">The URL.</param>
        /// <returns></returns>
        private static Image GetImageFromURL(string url)
        {
            try
            {
                HttpWebRequest httpWebRequest = (HttpWebRequest)HttpWebRequest.Create(url);
                HttpWebResponse httpWebReponse = (HttpWebResponse)httpWebRequest.GetResponse();
                Stream stream = httpWebReponse.GetResponseStream();
                return Image.FromStream(stream);
            }
            catch
            {
                return null;
            }
        }

    }
}
