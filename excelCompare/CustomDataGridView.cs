using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace excelCompare
{
    class CustomDataGridView : DataGridView
    {
        private int verticalWheelDelta;

        public ScrollBar GetVerticalScrollBar()
        {
            return VerticalScrollBar;
        }

        public ScrollBar GetHorizontalScrollBar()
        {
            return HorizontalScrollBar;
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
            if ((Control.ModifierKeys & (Keys.Shift | Keys.Alt)) == Keys.None && Control.MouseButtons == MouseButtons.None)
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
                                int finalValue = FirstDisplayedScrollingRowIndex - deltaRows;
                                if (finalValue < 0)
                                {
                                    finalValue = 0;
                                    verticalWheelDelta = 0;
                                }
                                else
                                {
                                    verticalWheelDelta -= (int)((float)deltaRows * (120f / (float)mouseWheelScrollLines));
                                }
                                if (finalValue >= 0 && finalValue < Rows.Count)
                                {
                                    FirstDisplayedScrollingRowIndex = finalValue;
                                }
                            }
                            else
                            {
                                int finalValue = FirstDisplayedScrollingRowIndex - deltaRows;
                                if (finalValue >= Rows.Count)
                                {
                                    finalValue = Rows.Count - 1;
                                    verticalWheelDelta = 0;
                                }
                                else
                                {
                                    verticalWheelDelta -= (int)((float)deltaRows * (120f / (float)mouseWheelScrollLines));
                                }
                                if (finalValue >= 0 && finalValue < Rows.Count)
                                {
                                    FirstDisplayedScrollingRowIndex = finalValue;
                                }
                            }
                        }
                    }
                }
            }
        }
    }
}
