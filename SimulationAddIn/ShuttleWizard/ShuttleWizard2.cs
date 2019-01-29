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
    public partial class ShuttleWizard2 : Form
    {
        ShuttleDataBL myShuttleData = null;
        public ShuttleWizard2()
        {
            InitializeComponent();
        }

        public ShuttleWizard2(ShuttleDataBL argShuttleData) : this()
        {
            myShuttleData = argShuttleData;
        }

        private void button3_Click(object sender, EventArgs e)
        {
            myShuttleData.SetArcFolder();
        }
    }
}
