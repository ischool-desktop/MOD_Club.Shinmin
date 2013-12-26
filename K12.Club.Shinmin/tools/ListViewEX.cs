using System;
using System.Collections.Generic;
using System.Text;
using System.ComponentModel;
using System.Windows.Forms;

namespace K12.Club.Shinmin
{
    /// <summary>
    /// ListView的DotNetBar行為正常化版本
    /// </summary>
    public class ListViewEX :
        DevComponents.DotNetBar.Controls.ListViewEx
    {
        private bool _HideSelection = false;
        private bool _FullRowSelect = false;

        /// <summary>
        /// 失去焦點時不顯視選取
        /// </summary>
        [Browsable(true)]
        public new bool HideSelection
        {
            get { return _HideSelection; }
            set { _HideSelection = value; }
        }

        /// <summary>
        /// 整行選取
        /// </summary>
        [Browsable(true)]
        public new bool FullRowSelect
        {
            get { return _FullRowSelect; }
            set
            {
                _FullRowSelect = value;
                base.FullRowSelect = _FullRowSelect;
            }
        }
        /// <summary>
        /// 建構子
        /// </summary>
        public ListViewEX()
        {
            base.FullRowSelect = true;
        }

        protected override void OnDrawSubItem(DrawListViewSubItemEventArgs e)
        {
            if ( !this.ContainsFocus && !_HideSelection )
            {
                ListViewItemStates status = e.Item.Selected ? ListViewItemStates.Selected : e.ItemState;
                base.OnDrawSubItem(new DrawListViewSubItemEventArgs(e.Graphics, e.Bounds, e.Item, e.SubItem, e.ItemIndex, e.ColumnIndex, e.Header, status));
            }
            else
                base.OnDrawSubItem(e);
        }
        protected override void OnDrawItem(DrawListViewItemEventArgs e)
        {
            base.FullRowSelect = _FullRowSelect;
            if ( !this.ContainsFocus && !_HideSelection )
            {
                ListViewItemStates status = e.Item.Selected ? ListViewItemStates.Selected : e.State;
                base.OnDrawItem(new DrawListViewItemEventArgs(e.Graphics, e.Item, e.Bounds, e.ItemIndex, status));
            }
            else
            {
                base.OnDrawItem(e);
            }
        }
        protected override void OnMouseDown(MouseEventArgs e)
        {
            if ( MultiSelect && e.Button == MouseButtons.Left )
                base.FullRowSelect = false;
            base.OnMouseDown(e);
        }
    }
}
