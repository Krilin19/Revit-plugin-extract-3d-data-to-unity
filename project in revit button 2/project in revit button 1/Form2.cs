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
    public partial class Form2 : Form
    {
        public Form2()
        {
            InitializeComponent();
        }
        //**************************************buttonsdisapear*********
        private void button1_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.OK;
            Close();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Close();
        }
        //**************************************buttonsdisapear*********


        public byte sliderValue()
        {

             
            Int32 max;
            max = this.trackBar4.Value;
            byte bytes3 = Convert.ToByte(max);
            return bytes3;

        }
        public byte sliderValue2()
        {


            Int32 max;
            max = this.trackBar5.Value;
            byte bytes3 = Convert.ToByte(max);
            return bytes3;

        }
        public byte sliderValue3()
        {


            Int32 max;
            max = this.trackBar6.Value;
            byte bytes3 = Convert.ToByte(max);
            return bytes3;

        }

        private void button3_Click_1(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.OK;
            Close();
        }

        private void button4_Click_1(object sender, EventArgs e)
        {
            Close();

        }

        private void label7_Click(object sender, EventArgs e)
        {

        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            //trackBar4.Value= textBox1.Text;
        }

        private void label10_Click(object sender, EventArgs e)
        {
            label1.Text = trackBar4.Value.ToString();
        }
    }
}
