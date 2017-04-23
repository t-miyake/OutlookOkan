using System.Collections.Generic;
using System.Windows.Forms;
using System.Drawing;

namespace OutlookOkan
{
    /// <summary>
    /// 項目の描画色を赤にする、CustomCheckedListBoxのカスタムコントロール
    /// </summary>
    internal class CustomCheckedListBox : CheckedListBox
    {
        /// <summary>
        /// 色を変えたい項目と同じIndexにtrueをaddすれば色が変わる。
        /// </summary>
        public List<bool> ColorFlag = new List<bool>();

        protected override void OnDrawItem(DrawItemEventArgs e)
        {
            var foreColor = Color.Black;

            if (ColorFlag.Count == 0 ? false : ColorFlag[e.Index])
            {
                foreColor = Color.Red;
            }

            var tweakedEventArgs = new DrawItemEventArgs(
                e.Graphics,
                e.Font,
                e.Bounds,
                e.Index,
                e.State,
                foreColor,
                e.BackColor);

            base.OnDrawItem(tweakedEventArgs);
        } 
    }
}