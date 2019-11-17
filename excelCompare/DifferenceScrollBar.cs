using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using System.Drawing.Drawing2D;

namespace excelCompare
{
    /// <summary>
    /// Enum for the scrollbar orientation.
    /// </summary>
    public enum ScrollBarOrientation
    {
        /// <summary>
        /// Indicates a horizontal scrollbar.
        /// </summary>
        Horizontal,

        /// <summary>
        /// Indicates a vertical scrollbar.
        /// </summary>
        Vertical
    }

    public enum DifferentMode
    {
        Same,
        Left,
        Right,
        Both
    }

    public interface IDifferenceAdapter
    {
        int DifferenceCount
        {
            get;
        }

        DifferentMode GetDifference(int index);

        float PageSizeRatio
        {
            get;
        }
    }

    public class DifferenceScrollBar : Control
    {
        private const int ARROW_WIDTH = 4;
        private const int FRAME_PADDING = 4;
        private const int DIFF_PADDING = FRAME_PADDING + 1;

        /// <summary>
        /// The value of the scrollbar.
        /// </summary>
        private int value;

        /// <summary>
        /// Selected row
        /// </summary>
        private int selection = -1;

        /// <summary>
        /// The scrollbar orientation - horizontal / vertical.
        /// </summary>
        private ScrollBarOrientation orientation = ScrollBarOrientation.Vertical;

        /// <summary>
        /// The scroll orientation in scroll events.
        /// </summary>
        private ScrollOrientation scrollOrientation = ScrollOrientation.VerticalScroll;

        /// <summary>
        /// Color for left difference indication
        /// </summary>
        private Color diffColor = Color.Red;

        /// <summary>
        /// Color for right difference indication
        /// </summary>
        private Color arrowColor = Color.Blue;

        /// <summary>
        /// The border color.
        /// </summary>
        private Color borderColor = Color.FromArgb(93, 140, 201);

        /// <summary>
        /// The border color in disabled state.
        /// </summary>
        private Color borderShadowColor = Color.Gray;

        /// <summary>
        /// Color for the 
        /// </summary>
        private Color thumbColor = Color.DarkGray;

        /// <summary>
        /// The progress timer for moving the thumb.
        /// </summary>
        private Timer progressTimer = new Timer();

        private IDifferenceAdapter dataSource = null;
        private int verticalWheelDelta = 0;

        [Category("Behavior")]
        [Description("Is raised, when the scrollbar was scrolled.")]
        public event ScrollEventHandler Scroll;

        /// <summary>
        /// Gets or sets the value.
        /// </summary>
        [Category("Behavior")]
        [Description("Gets or sets the current value.")]
        [DefaultValue(0)]
        public int Value
        {
            get
            {
                return this.value;
            }

            set
            {
                // no change or invalid value - return
                if (this.value == value || value < 0 || value >= GetDataCount())
                {
                    return;
                }

                this.value = value;

                // raise scroll event
                this.OnScroll(new ScrollEventArgs(ScrollEventType.ThumbPosition, -1, this.value, this.scrollOrientation));

                this.Refresh();
            }
        }

        [Category("Behavior")]
        [Description("Gets or sets the current selection.")]
        [DefaultValue(0)]
        public int Selection
        {
            get
            {
                return this.selection;
            }

            set
            {
                // no change or invalid value - return
                if (this.selection == value || value < 0 || value >= GetDataCount())
                {
                    return;
                }

                this.selection = value;

                this.Refresh();
            }
        }

        /// <summary>
        /// Gets or sets the border color.
        /// </summary>
        [Category("Appearance")]
        [Description("Gets or sets the difference color.")]
        [DefaultValue(typeof(Color), "Red")]
        public Color DiffColor
        {
            get
            {
                return this.diffColor;
            }

            set
            {
                this.diffColor = value;

                this.Invalidate();
            }
        }

        /// <summary>
        /// Gets or sets the border color.
        /// </summary>
        [Category("Appearance")]
        [Description("Gets or sets the arrow color.")]
        [DefaultValue(typeof(Color), "Blue")]
        public Color ArrowColor
        {
            get
            {
                return this.arrowColor;
            }

            set
            {
                this.arrowColor = value;

                this.Invalidate();
            }
        }

        /// <summary>
        /// Gets or sets the border color.
        /// </summary>
        [Category("Appearance")]
        [Description("Gets or sets the border color.")]
        [DefaultValue(typeof(Color), "93, 140, 201")]
        public Color BorderColor
        {
            get
            {
                return this.borderColor;
            }

            set
            {
                this.borderColor = value;

                this.Invalidate();
            }
        }

        /// <summary>
        /// Gets or sets the border shadow color.
        /// </summary>
        [Category("Appearance")]
        [Description("Gets or sets the border shadow color")]
        [DefaultValue(typeof(Color), "Gray")]
        public Color BorderShadowColor
        {
            get
            {
                return this.borderShadowColor;
            }

            set
            {
                this.borderShadowColor = value;

                this.Invalidate();
            }
        }

        /// <summary>
        /// Gets or sets the border color in disabled state.
        /// </summary>
        [Category("Appearance")]
        [Description("Gets or sets the thumb color.")]
        [DefaultValue(typeof(Color), "DarkGray")]
        public Color ThumbColor
        {
            get
            {
                return this.thumbColor;
            }

            set
            {
                this.thumbColor = value;

                this.Invalidate();
            }
        }

        public IDifferenceAdapter DataSource
        {
            get
            {
                return dataSource;
            }

            set
            {
                dataSource = value;
                ForceUpdate();
            }
        }

        private Bitmap differenceImage;
        public DifferenceScrollBar()
        {
            SetStyle(
               ControlStyles.OptimizedDoubleBuffer | ControlStyles.ResizeRedraw
               | ControlStyles.Selectable | ControlStyles.AllPaintingInWmPaint
               | ControlStyles.UserPaint, true);

            // timer for clicking and holding the mouse button
            // over/below the thumb and on the arrow buttons
            this.progressTimer.Interval = 20;
            this.progressTimer.Tick += this.ProgressTimerTick;
        }

        /// <summary>
        /// Handles the updating of the thumb.
        /// </summary>
        /// <param name="sender">The sender.</param>
        /// <param name="e">An object that contains the event data.</param>
        private void ProgressTimerTick(object sender, EventArgs e)
        {
            this.ProgressMovement();
        }

        Rectangle GetFrameRectangle()
        {
            Size clientSize = GetClientSize();
            int frameLeft = ClientRectangle.Left + ARROW_WIDTH;
            int frameRight = ClientRectangle.Right - 3;

            int frameTop = FRAME_PADDING;
            int frameHeight = GetThumbSize();

            int frameBottom = clientSize.Height  - FRAME_PADDING;
            if (GetDataCount() > 0)
            {
                frameTop = FRAME_PADDING + (int)((clientSize.Height - 2 * DIFF_PADDING) * (float)value / (float)GetDataCount());
                if (frameTop + frameHeight >= frameBottom)
                {
                    frameTop = frameBottom - frameHeight;
                }
            }
            return new Rectangle(frameLeft, frameTop, frameRight - frameLeft, frameHeight);
        }

        Rectangle GetDiffRectangle()
        {
            Size clientSize = GetClientSize();
            int frameLeft = ClientRectangle.Left + ARROW_WIDTH;
            int frameRight = ClientRectangle.Right - 3;

            int diffLeft = frameLeft + 5;
            int diffRight = frameRight - 5;

            return new Rectangle(diffLeft, DIFF_PADDING, diffRight - diffLeft + 1, clientSize.Height - 2 * DIFF_PADDING);
        }

        private Size GetClientSize()
        {
            Size size = new Size(ClientRectangle.Width, 2 * DIFF_PADDING + 1);
            int dataCount = GetDataCount();
            if (dataCount > 0)
            {
                int minHeight = ClientRectangle.Height - 2 * DIFF_PADDING;
                if (dataCount < minHeight)
                {
                    size.Height = dataCount + 2 * DIFF_PADDING;
                }
                else
                {
                    size.Height = ClientRectangle.Height;
                }
            }
            return size;
        }

        protected void RefreshDifferenceImage()
        {
            Rectangle diffRectangle = GetDiffRectangle();

            differenceImage = new Bitmap(diffRectangle.Width, diffRectangle.Height);
            Graphics graphics = Graphics.FromImage(differenceImage);
            graphics.SmoothingMode = SmoothingMode.None;
            graphics.InterpolationMode = InterpolationMode.NearestNeighbor;
            graphics.PixelOffsetMode = PixelOffsetMode.None;

            using (SolidBrush brush = new SolidBrush(this.BackColor))
            {
                graphics.FillRectangle(brush, new Rectangle(0, 0, differenceImage.Width, differenceImage.Height));
            }
            int dataCount = GetDataCount();
            if (dataCount > 0)
            {
                float lineSize = (1 / (float)dataCount) * differenceImage.Height;

                int diffMiddle = differenceImage.Width / 2;
                int left = 0;
                int right = differenceImage.Width - 1;

                using (Pen pen = new Pen(diffColor, lineSize))
                {
                    for (int i = 0; i < dataCount; i++)
                    {
                        float y = lineSize * i;
                        DifferentMode mode = dataSource.GetDifference(i);
                        if (mode == DifferentMode.Left || mode == DifferentMode.Both)
                        {
                            graphics.DrawLine(pen, left, y, diffMiddle - 1, y);
                        }
                        if (mode == DifferentMode.Right || mode == DifferentMode.Both)
                        {
                            graphics.DrawLine(pen, diffMiddle, y, right, y);
                        }
                    }
                }
            }
            graphics.Flush();
        }

        /// <summary>
        /// Raises the <see cref="Scroll"/> event.
        /// </summary>
        /// <param name="e">The <see cref="ScrollEventArgs"/> that contains the event data.</param>
        protected virtual void OnScroll(ScrollEventArgs e)
        {
            // if event handler is attached - raise scroll event
            if (this.Scroll != null)
            {
                this.Scroll(this, e);
            }
        }

        /// <summary>
        /// Paints the background of the control.
        /// </summary>
        /// <param name="e">A <see cref="PaintEventArgs"/> that contains information about the control to paint.</param>
        protected override void OnPaintBackground(PaintEventArgs e)
        {
            // no painting here
        }

        /// <summary>
        /// Paints the control.
        /// </summary>
        /// <param name="e">A <see cref="PaintEventArgs"/> that contains information about the control to paint.</param>
        protected override void OnPaint(PaintEventArgs e)
        {
            // sets the smoothing mode to none
            e.Graphics.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.None;

            // save client rectangle
            Rectangle rect = ClientRectangle;
            using (SolidBrush brush = new SolidBrush(this.BackColor))
            {
                e.Graphics.FillRectangle(brush, rect);
            }

            Rectangle diffRectangle = GetDiffRectangle();

            if (differenceImage == null)
            {
                RefreshDifferenceImage();
            }
            e.Graphics.DrawImage(differenceImage, diffRectangle, new Rectangle(0, 0, differenceImage.Width, differenceImage.Height), GraphicsUnit.Pixel);

            // Inner frame
            using (Pen pen = new Pen(borderColor, 1))
            {
                e.Graphics.DrawLine(pen, diffRectangle.Left, diffRectangle.Top - 1, diffRectangle.Right, diffRectangle.Top - 1);
                e.Graphics.DrawLine(pen, diffRectangle.Left, diffRectangle.Top - 1, diffRectangle.Left, diffRectangle.Bottom + 1);
            }
            using (Pen pen = new Pen(borderShadowColor, 1))
            {
                e.Graphics.DrawLine(pen, diffRectangle.Left, diffRectangle.Bottom + 1, diffRectangle.Right, diffRectangle.Bottom + 1);
                e.Graphics.DrawLine(pen, diffRectangle.Right, diffRectangle.Top - 1, diffRectangle.Right, diffRectangle.Bottom + 1);
            }

            // Thumb frame
            Color newColor = Color.FromArgb(128, thumbColor);
            Rectangle frameRectangle = GetFrameRectangle();
            using (Pen pen = new Pen(newColor, 1))
            {
                e.Graphics.DrawRectangle(pen, frameRectangle.Left, frameRectangle.Top, frameRectangle.Width, frameRectangle.Height);
            }

            if (selection != -1)
            {
                using (SolidBrush brush = new SolidBrush(arrowColor))
                {
                    float y = 0;
                    if (GetDataCount() > 0)
                    {
                        y = ((float)diffRectangle.Height * (float)selection / (float)GetDataCount());
                    }
                    PointF[] points = new PointF[]
                    {
                        new PointF(0, y + 1),
                        new PointF(4, y + DIFF_PADDING),
                        new PointF(0, y + 9)
                    };
                    e.Graphics.SmoothingMode = SmoothingMode.AntiAlias;
                    e.Graphics.InterpolationMode = InterpolationMode.HighQualityBicubic;
                    e.Graphics.PixelOffsetMode = PixelOffsetMode.HighQuality;
                    e.Graphics.FillPolygon(brush, points);
                }
            }
        }

        /// <summary>
        /// Enables the timer.
        /// </summary>
        private void EnableTimer()
        {
            // if timer is not already enabled - enable it
            if (!this.progressTimer.Enabled)
            {
                this.progressTimer.Interval = 600;
                this.progressTimer.Start();
            }
            else
            {
                // if already enabled, change tick time
                this.progressTimer.Interval = 10;
            }
        }

        /// <summary>
        /// Stops the progress timer.
        /// </summary>
        private void StopTimer()
        {
            this.progressTimer.Stop();
        }

        private void ProgressMovement()
        {
            Point mouseLocation = Point.Empty;
            GetCursorPos(ref mouseLocation);
            mouseLocation = PointToClient(mouseLocation);
            Size clientSize = GetClientSize();
            float ratio = (float)(mouseLocation.Y - 1) / (float)(clientSize.Height - 2);
            if (ratio > 1)
            {
                ratio = 1;
            }
            if (GetDataCount() > 0)
            {
                int newValue = (int)(ratio * GetDataCount());
                if ( newValue < 0 )
                {
                    newValue = 0;
                }
                if ( newValue > GetDataCount() )
                {
                    newValue = GetDataCount();
                }
                this.Value = newValue;
            }
        }

        /// <summary>
        /// Raises the MouseDown event.
        /// </summary>
        /// <param name="e">A <see cref="MouseEventArgs"/> that contains the event data.</param>
        protected override void OnMouseDown(MouseEventArgs e)
        {
            base.OnMouseDown(e);

            this.Focus();

            if (e.Button == MouseButtons.Left)
            {
                // prevents showing the context menu if pressing the right mouse
                // button while holding the left
                this.ContextMenuStrip = null;

                ProgressMovement();
                EnableTimer();
            }
        }

        /// <summary>
        /// Raises the MouseUp event.
        /// </summary>
        /// <param name="e">A <see cref="MouseEventArgs"/> that contains the event data.</param>
        protected override void OnMouseUp(MouseEventArgs e)
        {
            base.OnMouseUp(e);
            StopTimer();
        }

        /// <summary>
        /// Raises the MouseEnter event.
        /// </summary>
        /// <param name="e">A <see cref="EventArgs"/> that contains the event data.</param>
        protected override void OnMouseEnter(EventArgs e)
        {
            base.OnMouseEnter(e);
            Invalidate();
        }

        /// <summary>
        /// Raises the MouseLeave event.
        /// </summary>
        /// <param name="e">A <see cref="EventArgs"/> that contains the event data.</param>
        protected override void OnMouseLeave(EventArgs e)
        {
            base.OnMouseLeave(e);

            this.ResetScrollStatus();
        }

        /// <summary>
        /// Raises the MouseMove event.
        /// </summary>
        /// <param name="e">A <see cref="MouseEventArgs"/> that contains the event data.</param>
        protected override void OnMouseMove(MouseEventArgs e)
        {
            base.OnMouseMove(e);

            // moving and holding the left mouse button
            if (e.Button == MouseButtons.Left)
            {
                ProgressMovement();
            }
            else if (!this.ClientRectangle.Contains(e.Location))
            {
                this.ResetScrollStatus();
            }
            else if (e.Button == MouseButtons.None) // only moving the mouse
            {
            }
        }

        /// <summary>
        /// Performs the work of setting the specified bounds of this control.
        /// </summary>
        /// <param name="x">The new x value of the control.</param>
        /// <param name="y">The new y value of the control.</param>
        /// <param name="width">The new width value of the control.</param>
        /// <param name="height">The new height value of the control.</param>
        /// <param name="specified">A bitwise combination of the <see cref="BoundsSpecified"/> values.</param>
        protected override void SetBoundsCore(int x, int y, int width, int height, BoundsSpecified specified)
        {
            base.SetBoundsCore(x, y, width, height, specified);

            if (this.DesignMode)
            {
                this.SetUpScrollBar();
            }
        }

        /// <summary>
        /// Raises the <see cref="System.Windows.Forms.Control.SizeChanged"/> event.
        /// </summary>
        /// <param name="e">An <see cref="EventArgs"/> that contains the event data.</param>
        protected override void OnSizeChanged(EventArgs e)
        {
            base.OnSizeChanged(e);
            this.SetUpScrollBar();
            differenceImage = null;
        }

        /// <summary>
        /// Raises the <see cref="System.Windows.Forms.Control.EnabledChanged"/> event.
        /// </summary>
        /// <param name="e">An <see cref="EventArgs"/> that contains the event data.</param>
        protected override void OnEnabledChanged(EventArgs e)
        {
            base.OnEnabledChanged(e);

            if (this.Enabled)
            {
            }
            else
            {
            }

            this.Refresh();
        }


        /// <summary>
        /// Sends a message.
        /// </summary>
        /// <param name="wnd">The handle of the control.</param>
        /// <param name="msg">The message as int.</param>
        /// <param name="param">param - true or false.</param>
        /// <param name="lparam">Additional parameter.</param>
        /// <returns>0 or error code.</returns>
        /// <remarks>Needed for sending the stop/start drawing of the control.</remarks>
        [DllImport("user32.dll")]
        private static extern int SendMessage(
           IntPtr wnd,
           int msg,
           bool param,
           int lparam);

        /// <summary>
        /// Get cursor position
        /// </summary>
        /// <param name="lpPoint"></param>
        /// <returns></returns>
        [DllImport("user32.dll")]
        static extern bool GetCursorPos(ref Point lpPoint);

        /// <summary>
        /// Sets up the scrollbar.
        /// </summary>
        private void SetUpScrollBar()
        {
            // if no drawing - return
            differenceImage = null;

            this.Refresh();
        }

        /// <summary>
        /// Resets the scroll status of the scrollbar.
        /// </summary>
        private void ResetScrollStatus()
        {
            // get current mouse position
            Point pos = this.PointToClient(Cursor.Position);
            this.StopTimer();

            this.Refresh();
        }

        /// <summary>
        /// Calculates the height of the thumb.
        /// </summary>
        /// <returns>The height of the thumb.</returns>
        private int GetThumbSize()
        {
            Size clientSize = GetClientSize();
            int trackSize =
            this.orientation == ScrollBarOrientation.Vertical ?
            clientSize.Height - 2 * FRAME_PADDING : clientSize.Width - 2 * FRAME_PADDING;

            if (GetDataCount() == 0)
            {
                return trackSize;
            }

            float newThumbSize = dataSource != null ? trackSize * dataSource.PageSizeRatio : 0;
            return Convert.ToInt32(Math.Min((float)trackSize, Math.Max(newThumbSize, 4f)));
        }

        private int GetDataCount()
        {
            if (dataSource != null)
            {
                return dataSource.DifferenceCount;
            }
            return 0;
        }

        public void ForceUpdate()
        {
            differenceImage = null;
            this.Refresh();
        }

        /// <summary>Raises the <see cref="E:System.Windows.Forms.Control.MouseWheel" /> event.</summary>
		/// <param name="e">A <see cref="T:System.Windows.Forms.MouseEventArgs" /> that contains the event data.</param>
		protected override void OnMouseWheel(MouseEventArgs e)
        {
            base.OnMouseWheel(e);
            HandledMouseEventArgs handledMouseEventArgs = e as HandledMouseEventArgs;
            if (handledMouseEventArgs != null && handledMouseEventArgs.Handled)
            {
                return;
            }
            if ( GetDataCount() == 0 )
            {
                return;
            }

            if ((Control.ModifierKeys & (Keys.Shift | Keys.Alt)) == Keys.None && Control.MouseButtons == MouseButtons.None)
            {
                bool flag = (Control.ModifierKeys & Keys.Control) == Keys.None;
                {
                    if (handledMouseEventArgs != null)
                    {
                        handledMouseEventArgs.Handled = true;
                    }
                    int mouseWheelScrollLines = SystemInformation.MouseWheelScrollLines;
                    if (mouseWheelScrollLines != 0)
                    {
                        verticalWheelDelta += e.Delta;
                        float movement = (float)verticalWheelDelta / 120f;
                        if (mouseWheelScrollLines != -1)
                        {
                            int deltaRows = (int)((float)mouseWheelScrollLines * movement);
                            if (deltaRows != 0)
                            {
                                if (deltaRows > 0)
                                {
                                    int finalValue = value - deltaRows;
                                    if (finalValue < 0)
                                    {
                                        finalValue = 0;
                                        verticalWheelDelta = 0;
                                    }
                                    else
                                    {
                                        verticalWheelDelta -= (int)((float)deltaRows * (120f / (float)mouseWheelScrollLines));
                                    }
                                    Value = finalValue;
                                }
                                else
                                {
                                    int finalValue = value - deltaRows;
                                    if (finalValue > GetDataCount())
                                    {
                                        finalValue = GetDataCount();
                                        verticalWheelDelta = 0;
                                    }
                                    else
                                    {
                                        verticalWheelDelta -= (int)((float)deltaRows * (120f / (float)mouseWheelScrollLines));
                                    }
                                    Value = finalValue;
                                }
                            }
                        }
                    }
                }
            }
        }
    }
}
