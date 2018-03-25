using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Text;
using System.Windows.Forms;
using System.Drawing.Drawing2D;

namespace NetOffice.DeveloperToolbox.Controls.Text
{
	public partial class Scroller : UserControl
	{
		/// <summary>
		/// String list.
		/// </summary>
		string[] m_text = new string[0];

		/// <summary>
		/// Offset for animation.
		/// </summary>
		int m_scrollingOffset = 0;

		/// <summary>
		/// Top part size of text in percents.
		/// </summary>
		int m_topPartSizePercent = 50;

		/// <summary>
		/// Font, which is used to draw.
		/// </summary>
		Font m_font = new Font("Arial", 20, FontStyle.Bold, GraphicsUnit.Pixel);

		/// <summary>
		/// Constructor
		/// </summary>
		public Scroller()
		{
			InitializeComponent();

			// Enables double buffering (to remove flickering) and enables user paint.
			SetStyle(ControlStyles.OptimizedDoubleBuffer | ControlStyles.UserPaint | ControlStyles.AllPaintingInWmPaint, true);
		}

		/// <summary>
		/// Text to scroll.
		/// </summary>
		public string TextToScroll
		{
			get
			{
				return string.Join("\n", m_text);
			}
			set
			{
				string buffer = value;

				// Splits text by "\n" symbol.
				m_text = buffer.Split(new char[1] { '\n' });
			}
		}

		/// <summary>
		/// Timer interval.
		/// </summary>
		public int Interval
		{
			get
			{
				return m_timer.Interval;
			}
			set
			{
				m_timer.Interval = value;
			}
		}

		/// <summary>
		/// Font, which is used to draw.
		/// </summary>
		public Font TextFont
		{
			get
			{
				return m_font;
			}

			set
			{
				m_font = value;
			}
		}

		/// <summary>
		/// Top part size of text in percents (of control width).
		/// </summary>
		public int TopPartSizePercent
		{
			get
			{
				return m_topPartSizePercent;
			}
			set
			{
				if ((value >= 10) && (value <= 100))
				{
					m_topPartSizePercent = value;
				}
				else
					throw new InvalidEnumArgumentException("The value must be more than zero. and less than 100.");
			}
		}

		/// <summary>
		/// Paint handler.
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		private void OnPaint(object sender, PaintEventArgs e)
		{
			// Sets antialiasing mode for better quality.
			e.Graphics.SmoothingMode = SmoothingMode.AntiAlias;

			// Prepares background.
			e.Graphics.FillRectangle(new SolidBrush(this.BackColor), this.ClientRectangle);

			// Creates GraphicsPath for text.
			GraphicsPath path = new GraphicsPath();

			// Visible lines counter;
			int visibleLines = 0;

			for (int i = m_text.Length - 1; i >= 0; i--)
			{

				Point pt = new Point((int)((this.ClientSize.Width - e.Graphics.MeasureString(m_text[i], m_font).Width) / 2),
					(int)(m_scrollingOffset + this.ClientSize.Height - (m_text.Length - i) * m_font.Size));

				// Adds visible lines to path.
				if ((pt.Y + this.Font.Size > 0) && (pt.Y < this.Height))
				{
					path.AddString(m_text[i], m_font.FontFamily, (int)m_font.Style, m_font.Size,
						pt, StringFormat.GenericTypographic);

					visibleLines++;
				}

			}

			// For repeat scrolling.
			if ((visibleLines == 0) && (m_scrollingOffset < 0))
			{
				m_scrollingOffset = (int)this.Font.SizeInPoints * m_text.Length;
			}

			int topSizeWidth = (int)(this.Width * m_topPartSizePercent / 100.0f);

			// Wraps Graphics path from rectangle to trapeze.
			path.Warp(
				new PointF[4]
				{
					new PointF((this.Width - topSizeWidth) / 2, 0),
					new PointF(this.Width - (this.Width - topSizeWidth) / 2, 0),
					new PointF(0, this.Height),
					new PointF(this.Width, this.Height)
				},
				new RectangleF(this.ClientRectangle.X, this.ClientRectangle.Y, this.ClientRectangle.Width, this.ClientRectangle.Height),
				null,
				WarpMode.Perspective
				);

			// Draws wrapped path.
			e.Graphics.FillPath(new SolidBrush(this.ForeColor), path);
			path.Dispose();

			// Draws fog effect with help of gradient brush with alpha colors.
			using (Brush br = new LinearGradientBrush(new Point(0, 0), new Point(0, this.Height),
				Color.FromArgb(255, this.BackColor), Color.FromArgb(0, this.BackColor)))
			{
				e.Graphics.FillRectangle(br, this.ClientRectangle);
			}
		}

		/// <summary>
		/// Starts the animation from the beginning.
		/// </summary>
		public void Start()
		{
			// Calculates scrolling offset.
			m_scrollingOffset = (int)this.Font.SizeInPoints * m_text.Length;
			m_timer.Start();
		}

		/// <summary>
		/// Stops the animation.
		/// </summary>
		public void Stop()
		{
			m_timer.Stop();
		}

		/// <summary>
		/// Timer handler.
		/// </summary>
		private void OnTimerTick(object sender, EventArgs e)
		{
			// Changes the offset.
			m_scrollingOffset--;

			// Repaints whole control area.
			Invalidate();
		}
	}
}
