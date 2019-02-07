using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace SimulationAddIn
{
    public partial class DHAMenu : Form
    {
        public DHAMenu()
        {
            InitializeComponent();
        }

        private void btnViewData_Click(object sender, EventArgs e)
        {
            DataViewer myDVForm = new DataViewer();
            this.Hide();
            myDVForm.Show();
        }
    }
}
