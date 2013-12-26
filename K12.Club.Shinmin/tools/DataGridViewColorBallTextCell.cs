using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Drawing.Drawing2D;
using System.Drawing;

namespace K12.Club.Shinmin
{
    class DataGridViewColorBallTextCell : DataGridViewTextBoxCell
    {
        private Color _Color = Color.Transparent;
        public Color Color
        {
            get { return _Color; }
            set { _Color = value; }
        }
        protected override void Paint(System.Drawing.Graphics graphics, System.Drawing.Rectangle clipBounds, System.Drawing.Rectangle cellBounds, int rowIndex, DataGridViewElementStates cellState, object value, object formattedValue, string errorText, DataGridViewCellStyle cellStyle, DataGridViewAdvancedBorderStyle advancedBorderStyle, DataGridViewPaintParts paintParts)
        {
            if (_Color != Color.Transparent)
            {
                SmoothingMode mode = graphics.SmoothingMode;//跟妳說喔，NotNetBar很機車喔，如果你把SmoothingMode改掉沒改回去，格線會亂亂劃喔。
                graphics.SmoothingMode = SmoothingMode.HighQuality;
                base.Paint(graphics, clipBounds, cellBounds, rowIndex, cellState, null, null, errorText, cellStyle, advancedBorderStyle, paintParts);

                int w = cellBounds.Height - 9,
                    x = cellBounds.X + 3,
                    y = cellBounds.Y + 4;
                Color[] myColors = { _Color, Color.White, _Color, _Color };
                float[] myPositions = { 0.0f, 0.05f, 0.6f, 1.0f };
                ColorBlend myBlend = new ColorBlend();
                myBlend.Colors = myColors;
                myBlend.Positions = myPositions;
                using (LinearGradientBrush brush = new LinearGradientBrush(new Point(x, y), new Point(x + w, y + w), Color.White, _Color))
                {
                    brush.InterpolationColors = myBlend;
                    brush.GammaCorrection = true;
                    graphics.FillRectangle(brush, x, y, w, w);
                }
                graphics.DrawRectangle(new Pen(Color.Black), x, y, w, w);

                cellBounds = new System.Drawing.Rectangle(cellBounds.X + cellBounds.Height - 4, cellBounds.Y, cellBounds.Width - cellBounds.Height + 4, cellBounds.Height);
                graphics.SmoothingMode = mode;
            }
            base.Paint(graphics, clipBounds, cellBounds, rowIndex, cellState, value, formattedValue, errorText, cellStyle, advancedBorderStyle, paintParts);
        }

        public override Type EditType
        {
            get
            {
                return null;
            }
        }
    }

    class DataGridViewColorBallTextColumn : DataGridViewColumn
    {
        public DataGridViewColorBallTextColumn()
        {
            this.CellTemplate = new DataGridViewColorBallTextCell();
        }
    }
}
