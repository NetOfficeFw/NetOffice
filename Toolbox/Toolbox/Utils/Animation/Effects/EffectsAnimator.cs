using System;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NetOffice.DeveloperToolbox.Utils.Animation.Effects
{
    internal enum EffectsKind
    {   
        FadeIn = 0,
        SlideBottomToTop =1,
        SlideTopToBottom = 2,
        SlideLeftToRight = 3,
        SlideRightToLeft = 4,
        Collapse = 5
    }

    internal static class EffectsAnimator
    {
        public enum Effect { Roll, Slide, Center, Blend }

        public static void Animate(Control ctl, Effect effect, int msec, int angle)
        {
            int flags = effmap[(int)effect];

            if (ctl.Visible)
            { 
                flags |= 0x10000;
                angle += 180;
            }
            else
            {
                if (ctl.TopLevelControl == ctl)
                    flags |= 0x20000;
                else if (effect == Effect.Blend)
                    throw new ArgumentException();
            }
            flags |= dirmap[(angle % 360) / 45];
             WinAPI.AnimateWindow(ctl.Handle, msec, flags);
            //if (!ok) throw new Exception("Animation failed");
            ctl.Visible = !ctl.Visible;
        }

        private static int[] dirmap = { 1, 5, 4, 6, 2, 10, 8, 9 };
        private static int[] effmap = { 0, 0x40000, 0x10, 0x80000 };


        internal static void DoEffect(Control ctrl, EffectsKind kind = EffectsKind.FadeIn, bool useSlideIfPossible = true, int animationSpeedInMS = 250)
        {
            //Animate(ctrl, Effect.Slide, 150, 180);
            //return;

            int flags = 0;

            switch (kind)
            {
                case EffectsKind.SlideBottomToTop:
                    flags = WinAPI.AW_ACTIVATE | WinAPI.AW_VER_NEGATIVE;
                    if (useSlideIfPossible)
                        flags |= WinAPI.AW_SLIDE;
                    break;
                case EffectsKind.SlideTopToBottom:
                    flags = WinAPI.AW_ACTIVATE|WinAPI.AW_VER_POSITIVE;
                    if (useSlideIfPossible)
				        flags |= WinAPI.AW_SLIDE;
                    break;
                case EffectsKind.SlideLeftToRight:
                    flags = WinAPI.AW_ACTIVATE|WinAPI.AW_HOR_POSITIVE;
                    if (useSlideIfPossible)
				        flags |= WinAPI.AW_SLIDE;
                    break;
                case EffectsKind.SlideRightToLeft:
                    flags = WinAPI.AW_ACTIVATE | WinAPI.AW_HOR_NEGATIVE;
                    if (useSlideIfPossible)
                        flags |= WinAPI.AW_SLIDE;
                    break;
                case EffectsKind.Collapse:
                    flags = WinAPI.AW_ACTIVATE | WinAPI.AW_CENTER;
                    break;
                default: // EffectsKind.FadeIn
                    flags = WinAPI.AW_ACTIVATE | WinAPI.AW_BLEND;
                    break;
            }

            WinAPI.AnimateWindow(ctrl.Handle, animationSpeedInMS, flags);
        }

        private class WinAPI
        {
            /// <summary>
            /// Animates the window from left to right. This flag can be used with roll or slide animation.
            /// </summary>
            public const int AW_HOR_POSITIVE = 0X1;
            /// <summary>
            /// Animates the window from right to left. This flag can be used with roll or slide animation.
            /// </summary>
            public const int AW_HOR_NEGATIVE = 0X2;
            /// <summary>
            /// Animates the window from top to bottom. This flag can be used with roll or slide animation.
            /// </summary>
            public const int AW_VER_POSITIVE = 0X4;
            /// <summary>
            /// Animates the window from bottom to top. This flag can be used with roll or slide animation.
            /// </summary>
            public const int AW_VER_NEGATIVE = 0X8;

            /// <summary>
            /// Makes the window appear to collapse inward if AW_HIDE is used or expand outward if the AW_HIDE is not used.
            /// </summary>
            public const int AW_CENTER = 0X10;
            /// <summary>
            /// Hides the window. By default, the window is shown.
            /// </summary>
            public const int AW_HIDE = 0x10000;
            /// <summary>
            /// Activates the window.
            /// </summary>
            public const int AW_ACTIVATE = 0X20000;
            /// <summary>
            /// Uses slide animation. By default, roll animation is used.
            /// </summary>
            public const int AW_SLIDE = 0X40000;
            /// <summary>
            /// Uses a fade effect. This flag can be used only if hwnd is a top-level window.
            /// </summary>
            public const int AW_BLEND = 0X80000;

            /// <summary>
            /// Animates a window.
            /// </summary>
            [DllImport("user32.dll", CharSet = CharSet.Auto)]
            public static extern int AnimateWindow(IntPtr hwand, int dwTime, int dwFlags);
        } 
    }
}
