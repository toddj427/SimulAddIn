using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace SimulationAddIn.ShuttleWizard
{
    public partial class ShuttleWizard1 : Form
    {
        ShuttleDataBL myShuttleData = new ShuttleDataBL();

        public ShuttleWizard1()
        {
            InitializeComponent();

        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void ShuttleWizard1_Load(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            ShuttleWizard2 myShuttleWizard2 = new ShuttleWizard2(myShuttleData);
            myShuttleWizard2.Show();

        }
    }
}
