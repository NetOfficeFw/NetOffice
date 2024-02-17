using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using System.Text;
using System.Drawing;
using System.Drawing.Drawing2D;

namespace Sample.Addin
{
    /// <summary>
    /// Pong, original developed from David Sosby - http://dragunsflame.cjb.net 
    /// </summary>
    public class Pong : Control
    {
        #region Properties

        Pen BallPen { get; set; }
        Pen PaddlePen { get; set; }
        int XVelocity { get; set; }
        int YVelocity { get; set; }
        int BallXPosition { get; set; }
        int BallYPosition { get; set; }
        Random RandomNumbers { get; set; }
        bool IsActive { get; set; }
        int p_x; Timer gameClock;

        #endregion

        #region Ctor

        public Pong()
        {
            XVelocity = -5;
            YVelocity = -5;
            RandomNumbers = new Random();
            BallPen = new Pen(Color.Red, 1);
            PaddlePen = new Pen(Color.Black, 5);
            this.MouseEnter += new EventHandler(Pong_MouseEnter);
            this.MouseLeave += new EventHandler(Pong_MouseLeave);
            this.MouseMove += new MouseEventHandler(Pong_MouseMove);
            this.Paint += new PaintEventHandler(Pong_Paint);
        }

        #endregion

        #region Trigger

        void Pong_MouseLeave(object sender, EventArgs e)
        {
            Cursor.Show();
        }

        void Pong_MouseEnter(object sender, EventArgs e)
        {
            Cursor.Hide();
        }

        private void Pong_MouseMove(Object sender, MouseEventArgs e)
        {
            p_x = e.X - 25;
        }
        
        private void Pong_Paint(Object sender, PaintEventArgs e)
        {
            DrawGame(e.Graphics);
        }

        #endregion

        #region Methods

        public void StartGame()
        {
            gameClock = new Timer();
            BallXPosition =  RandomNumbers.Next(this.Width - 50) + 10;
            BallYPosition = RandomNumbers.Next(this.Height - 50) + 10;
            p_x = 100;
            IsActive = true;
            gameClock.Interval = 15;
            gameClock.Tick += new EventHandler(GameLoop);
            gameClock.Start();
        }

        private void DrawGame(Graphics g)
        {
            if (IsActive)
            {
                g.DrawEllipse(BallPen, BallXPosition, BallYPosition, 5, 5);
                g.DrawLine(PaddlePen, p_x, this.Height-5, p_x + 50, this.Height-5);
            }
        }

        private void GameLoop(Object sender, EventArgs e)
        {
            if (BallXPosition <= 0)
                XVelocity *= -1;
            if (BallYPosition <= 0)
                YVelocity *= -1;

            if (BallXPosition >= this.Width)
                XVelocity *= 1;

            if (BallXPosition + 5 >= this.Width)
                XVelocity = -1 + (RandomNumbers.Next(2) - 2);


            if (BallYPosition + 5 >= this.Height && BallYPosition + 5 <= this.Height + 5 && BallXPosition + 3 >= p_x && BallXPosition + 3 <= p_x + 50)
                YVelocity *= -1;

            if (BallYPosition >= this.Height || BallXPosition >= this.Width)
                DoGameOver();

            BallXPosition += XVelocity;
            BallYPosition += YVelocity;

            Invalidate();
        }

        private void DoGameOver()
        {
            gameClock.Stop();
            gameClock.Dispose();
            IsActive = false;
            Invalidate();
        }

        #endregion
    }

}
