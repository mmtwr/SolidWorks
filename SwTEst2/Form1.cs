using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Macro1;
using Macro1.csproj;

namespace SwTEst2
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            
        }

        private void Macro1_but_Click(object sender, EventArgs e)
        {
            SolidWorksMacro s = new SolidWorksMacro();
            s.Macro1();         
        }

        private void Macro2_but_Click(object sender, EventArgs e)
        {
            SolidWorksMacro2 s2 = new SolidWorksMacro2();
            s2.Macro2();
        }
    }
}
