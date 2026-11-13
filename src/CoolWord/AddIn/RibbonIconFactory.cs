using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Text;

namespace CoolWord.AddIn
{
    /// <summary>
    /// Nhà máy tạo biểu tượng Ribbon độ nét cao (32x32) cho các menu tích hợp:
    /// 1. Gửi Mail hàng loạt (Outlook + Gmail)
    /// 2. Ảnh (Khung ảnh + Mũi tên xanh)
    /// 3. Google Dịch (Logo Google Translate xanh/trắng)
    /// </summary>
    public static class RibbonIconFactory
    {
        private static Bitmap _mailIcon;
        private static Bitmap _imageIcon;
        private static Bitmap _translateIcon;

        private static readonly object LockObj = new object();

        public static Bitmap CreateMailMergeIcon()
        {
            lock (LockObj)
            {
                if (_mailIcon != null) return _mailIcon;

                Bitmap bmp = new Bitmap(32, 32);
                using (Graphics g = Graphics.FromImage(bmp))
                {
                    g.SmoothingMode = SmoothingMode.AntiAlias;
                    g.InterpolationMode = InterpolationMode.HighQualityBicubic;
                    g.TextRenderingHint = TextRenderingHint.ClearTypeGridFit;
                    g.Clear(Color.Transparent);

                    // 1. Phía trái: Biểu tượng Outlook (Màu xanh dương đậm #0078D4)
                    Rectangle outRect = new Rectangle(2, 4, 15, 23);
                    using (GraphicsPath path = CreateRoundedRectangle(outRect, 3))
                    using (LinearGradientBrush brush = new LinearGradientBrush(outRect, Color.FromArgb(0, 120, 215), Color.FromArgb(0, 90, 170), LinearGradientMode.Vertical))
                    {
                        g.FillPath(brush, path);
                    }

                    // Chữ O màu trắng
                    using (Font fontO = new Font("Segoe UI", 11, FontStyle.Bold))
                    using (Brush whiteBrush = new SolidBrush(Color.White))
                    {
                        StringFormat sf = new StringFormat
                        {
                            Alignment = StringAlignment.Center,
                            LineAlignment = StringAlignment.Center
                        };
                        g.DrawString("O", fontO, whiteBrush, new RectangleF(2, 4, 15, 23), sf);
                    }

                    // 2. Phía phải đè lên nhẹ: Phong bì Gmail (Nền trắng, viền đỏ #EA4335)
                    Rectangle gmRect = new Rectangle(13, 8, 17, 19);
                    using (GraphicsPath pathG = CreateRoundedRectangle(gmRect, 2))
                    {
                        using (SolidBrush bg = new SolidBrush(Color.White))
                        {
                            g.FillPath(bg, pathG);
                        }
                        using (Pen pen = new Pen(Color.FromArgb(220, 220, 220), 1))
                        {
                            g.DrawPath(pen, pathG);
                        }
                    }

                    // Nắp gập chữ M màu đỏ của Gmail
                    using (Pen redPen = new Pen(Color.FromArgb(234, 67, 53), 2.2f))
                    {
                        redPen.StartCap = LineCap.Round;
                        redPen.EndCap = LineCap.Round;
                        redPen.LineJoin = LineJoin.Round;

                        Point p1 = new Point(14, 10);
                        Point p2 = new Point(21, 17);
                        Point p3 = new Point(28, 10);

                        g.DrawLine(redPen, p1, p2);
                        g.DrawLine(redPen, p2, p3);
                        g.DrawLine(redPen, 14, 10, 14, 25);
                        g.DrawLine(redPen, 28, 10, 28, 25);
                    }
                }

                _mailIcon = bmp;
                return _mailIcon;
            }
        }

        public static Bitmap CreateImageToolsIcon()
        {
            lock (LockObj)
            {
                if (_imageIcon != null) return _imageIcon;

                Bitmap bmp = new Bitmap(32, 32);
                using (Graphics g = Graphics.FromImage(bmp))
                {
                    g.SmoothingMode = SmoothingMode.AntiAlias;
                    g.Clear(Color.Transparent);

                    // 1. Khung tranh (Picture Frame)
                    Rectangle frameRect = new Rectangle(3, 4, 22, 18);
                    using (GraphicsPath framePath = CreateRoundedRectangle(frameRect, 2))
                    {
                        using (SolidBrush bgBrush = new SolidBrush(Color.FromArgb(250, 250, 252)))
                        {
                            g.FillPath(bgBrush, framePath);
                        }
                        using (Pen borderPen = new Pen(Color.FromArgb(80, 80, 80), 1.8f))
                        {
                            g.DrawPath(borderPen, framePath);
                        }
                    }

                    // Mặt trời nhỏ (vàng/cam)
                    using (SolidBrush sunBrush = new SolidBrush(Color.FromArgb(245, 166, 35)))
                    {
                        g.FillEllipse(sunBrush, 18, 7, 4, 4);
                    }

                    // Núi (Triangles)
                    using (GraphicsPath mountainPath = new GraphicsPath())
                    {
                        mountainPath.AddPolygon(new Point[]
                        {
                            new Point(4, 21),
                            new Point(11, 12),
                            new Point(18, 21)
                        });
                        mountainPath.AddPolygon(new Point[]
                        {
                            new Point(14, 21),
                            new Point(19, 14),
                            new Point(24, 21)
                        });

                        using (SolidBrush mBrush = new SolidBrush(Color.FromArgb(100, 105, 115)))
                        {
                            g.FillPath(mBrush, mountainPath);
                        }
                    }

                    // 2. Mũi tên xanh dương trích xuất hướng sang phải (Blue Arrow ->)
                    using (Pen arrowPen = new Pen(Color.FromArgb(0, 130, 230), 2.4f))
                    {
                        arrowPen.StartCap = LineCap.Round;
                        arrowPen.EndCap = LineCap.Round;

                        Point startPt = new Point(16, 25);
                        Point endPt = new Point(29, 25);
                        g.DrawLine(arrowPen, startPt, endPt);

                        // Đầu mũi tên >
                        g.DrawLine(arrowPen, 24, 20, 29, 25);
                        g.DrawLine(arrowPen, 24, 30, 29, 25);
                    }
                }

                _imageIcon = bmp;
                return _imageIcon;
            }
        }

        public static Bitmap CreateGoogleTranslateIcon()
        {
            lock (LockObj)
            {
                if (_translateIcon != null) return _translateIcon;

                Bitmap bmp = new Bitmap(32, 32);
                using (Graphics g = Graphics.FromImage(bmp))
                {
                    g.SmoothingMode = SmoothingMode.AntiAlias;
                    g.TextRenderingHint = TextRenderingHint.ClearTypeGridFit;
                    g.Clear(Color.Transparent);

                    // 1. Khối nền xám/trắng phía sau (Ký tự Á/Âu/Á Đông)
                    Rectangle backRect = new Rectangle(11, 10, 18, 18);
                    using (GraphicsPath pathBack = CreateRoundedRectangle(backRect, 3))
                    {
                        using (SolidBrush backBrush = new SolidBrush(Color.FromArgb(235, 238, 242)))
                        {
                            g.FillPath(backBrush, pathBack);
                        }
                        using (Pen pen = new Pen(Color.FromArgb(200, 205, 215), 1))
                        {
                            g.DrawPath(pen, pathBack);
                        }
                    }

                    // Chữ 文 (Văn / Ngôn ngữ)
                    using (Font fontZh = new Font("SimSun", 10, FontStyle.Regular))
                    using (Brush zhBrush = new SolidBrush(Color.FromArgb(100, 105, 115)))
                    {
                        StringFormat sf = new StringFormat
                        {
                            Alignment = StringAlignment.Center,
                            LineAlignment = StringAlignment.Center
                        };
                        g.DrawString("文", fontZh, zhBrush, new RectangleF(13, 11, 16, 17), sf);
                    }

                    // 2. Khối xanh Google phía trước (Màu xanh Google #4285F4)
                    Rectangle frontRect = new Rectangle(3, 4, 18, 18);
                    using (GraphicsPath pathFront = CreateRoundedRectangle(frontRect, 3))
                    {
                        using (LinearGradientBrush blueBrush = new LinearGradientBrush(
                            frontRect,
                            Color.FromArgb(66, 133, 244),
                            Color.FromArgb(26, 115, 232),
                            LinearGradientMode.Vertical))
                        {
                            g.FillPath(blueBrush, pathFront);
                        }
                    }

                    // Nếp gấp origami góc dưới bên phải khối xanh
                    Point[] fold = new Point[]
                    {
                        new Point(15, 22),
                        new Point(21, 16),
                        new Point(15, 16)
                    };
                    using (SolidBrush foldBrush = new SolidBrush(Color.FromArgb(24, 90, 188)))
                    {
                        g.FillPolygon(foldBrush, fold);
                    }

                    // Chữ G màu trắng
                    using (Font fontG = new Font("Product Sans", 10, FontStyle.Bold))
                    using (Brush whiteBrush = new SolidBrush(Color.White))
                    {
                        StringFormat sf = new StringFormat
                        {
                            Alignment = StringAlignment.Center,
                            LineAlignment = StringAlignment.Center
                        };
                        g.DrawString("G", fontG, whiteBrush, new RectangleF(3, 4, 16, 18), sf);
                    }
                }

                _translateIcon = bmp;
                return _translateIcon;
            }
        }

        private static GraphicsPath CreateRoundedRectangle(Rectangle rect, int radius)
        {
            GraphicsPath path = new GraphicsPath();
            int diameter = radius * 2;
            Rectangle arc = new Rectangle(rect.Location, new Size(diameter, diameter));

            path.AddArc(arc, 180, 90);
            arc.X = rect.Right - diameter;
            path.AddArc(arc, 270, 90);
            arc.Y = rect.Bottom - diameter;
            path.AddArc(arc, 0, 90);
            arc.X = rect.Left;
            path.AddArc(arc, 90, 90);
            path.CloseFigure();
            return path;
        }
    }
}
