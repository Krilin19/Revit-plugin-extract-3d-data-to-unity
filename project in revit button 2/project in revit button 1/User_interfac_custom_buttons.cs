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
    public partial class User_interfac_custom_buttons : Form
    {
        public User_interfac_custom_buttons()
        {
            InitializeComponent();
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void button2_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Close();
        }

        public double getDistance()
        {
            double d;
            Double.TryParse(textBox1.Text, out d);
            return d;

        

        }

        public bool ishorizontal()
        {
            if (vertical.Checked)
                return true;
            else
                return false;
        }



        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }
    }
}
