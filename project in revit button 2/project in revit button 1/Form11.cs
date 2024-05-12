using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace AJC_Commands_1
{
    public partial class Form11 : Form
    {
        public Form11()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.OK;
            Close();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            foreach (var item in listBox1.SelectedItems)
            {
                if (!listBox2.Items.Contains(item))
                {
                    listBox2.Items.Add(item);
                }
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            List<object> objs = new List<object>();

            foreach (var item in listBox2.SelectedItems)
            {
                objs.Add(item);

            }
            foreach (var item in objs)
            {
                listBox2.Items.Remove(item);
            }
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }
    }
}
