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
    // DHA menu is the main form used to select the next step of events
    // Navigate between sheets
    // Run the model
    // Write the input files for the model
    // Manage Scenarios
    public partial class DHAMenu : Form
    {
        public DHAMenu()
        {
            InitializeComponent();
        }

        // Viewing data switches to the data viewer form to navigate between workbooks
        private void btnViewData_Click(object sender, EventArgs e)
        {
            DataViewer myDVForm = new DataViewer();
            this.Hide();
            myDVForm.Show();
        }

        // TODO: Add Run the Model
        // TODO: Add Scenarios
    }
}
