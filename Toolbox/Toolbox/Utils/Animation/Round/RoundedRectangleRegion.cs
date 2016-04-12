using System;
using System.Drawing;
using System.Drawing.Drawing2D;

namespace NetOffice.DeveloperToolbox.Utils.Animation.Round
{
    public class RoundedRectangleRegion
    {
        public Region GetRoundedRect(RectangleF baseRect, int radius)
        {
            if (radius <= 0)
                return new Region(baseRect);

            if (radius >= (Math.Min(baseRect.Width, baseRect.Height) / 2.0))
                return GetCapsule(baseRect);

            float diameter = radius + radius;
            RectangleF arcRect = new RectangleF(baseRect.Location, new SizeF(diameter, diameter));
            GraphicsPath rr = new GraphicsPath();

            rr.AddArc(arcRect, 180, 90);
            arcRect.X = baseRect.Right - diameter;
            rr.AddArc(arcRect, 270, 90);
            arcRect.Y = baseRect.Bottom - diameter;
            rr.AddArc(arcRect, 0, 90);
            arcRect.X = baseRect.Left;
            rr.AddArc(arcRect, 90, 90);
            rr.CloseFigure();

            return new Region(rr);
        }

        public Region GetCapsule(RectangleF baseRect)
        {
            float diameter;
            RectangleF arcRect;
            GraphicsPath rr = new GraphicsPath();

            try
            {
                if (baseRect.Width > baseRect.Height)
                {
                    diameter = baseRect.Height;
                    arcRect = new RectangleF(baseRect.Location, new SizeF(diameter, diameter));
                    rr.AddArc(arcRect, 90, 180);
                    arcRect.X = baseRect.Right - diameter;
                    rr.AddArc(arcRect, 270, 180);
                }
                else if (baseRect.Height > baseRect.Width)
                {
                    diameter = baseRect.Width;
                    arcRect = new RectangleF(baseRect.Location, new SizeF(diameter, diameter));
                    rr.AddArc(arcRect, 180, 180);
                    arcRect.Y = baseRect.Bottom - diameter;
                    rr.AddArc(arcRect, 0, 180);
                }
                else
                {
                    rr.AddEllipse(baseRect);
                }
            }
            catch (Exception e)
            {
                string sLastError = e.Message;
                rr.AddEllipse(baseRect);
            }
            finally
            {
                rr.CloseFigure();
            }
            return new Region(rr);
        }
    }
}
