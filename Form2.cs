using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Serial2
{
    public partial class Form2 : Form
    {
        public Form2()
        {
            InitializeComponent();
        }
        private Double Form2_value;

        public Double Passvalue
        {
            get { return this.Form2_value; }
            set { this.Form2_value = value; }  // 다른폼(Form1)에서 전달받은 값을 쓰기
        }

        private void Form2_Load(object sender, EventArgs e)
        {
            numericUpDown1.Value = Convert.ToDecimal(Passvalue); // 다른폼(Form1)에서 전달받은 값을 변수에 저장
            numericUpDown1.DecimalPlaces = 2;
        }

    }
}
