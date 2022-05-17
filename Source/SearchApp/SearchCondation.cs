using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace SearchApp
{
    public partial class SearchCondation : Form
    {
        public Condation ResultCondation = null;
        public SearchCondation()
        {
            InitializeComponent();
        }

        private void btSearch_Click(object sender, EventArgs e)
        {
            ResultCondation = null;

            if (string.IsNullOrEmpty(tbTuGang.Text.Trim())
                && string.IsNullOrEmpty(tbTuLei.Text.Trim())
                && string.IsNullOrEmpty(tbTuXi.Text.Trim())
                && string.IsNullOrEmpty(tbTuZu.Text.Trim())
                && string.IsNullOrEmpty(tbYaGang.Text.Trim())
                && string.IsNullOrEmpty(tbYaLei.Text.Trim())
                )
            {
                MessageBox.Show("查询条件不能全部为空");
            }
            else
            {
                ResultCondation = new Condation();
                ResultCondation.TuGang = tbTuGang.Text;
                ResultCondation.TuLei = tbTuLei.Text;
                ResultCondation.TuXi = tbTuXi.Text;
                ResultCondation.TuZu = tbTuZu.Text;
                ResultCondation.YaGang = tbYaGang.Text;
                ResultCondation.YaLei = tbYaLei.Text;
                this.DialogResult = DialogResult.OK;
            }
        }

        private void btClear_Click(object sender, EventArgs e)
        {
            tbTuGang.Clear();
            tbTuLei.Clear();
            tbTuXi.Clear();
            tbTuZu.Clear();
            tbYaGang.Clear();
            tbYaLei.Clear();
        }
    }
}
