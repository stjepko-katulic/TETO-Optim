using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace TETOAplikacija
{
    public partial class Form_Opcije : Form
    {
        public Form_Opcije()
        {
            InitializeComponent();
        }

        private void proba(object sender, FormClosingEventArgs e)
        {
            this.Hide();
            e.Cancel = true; //ova naredba cancelira event
        }

        // ova metoda (event) sprijecava da se text u comboboxu editira
        private void textedit(object sender, KeyPressEventArgs e)
        {
            e.Handled = true;
        }
    
        
        
    }
}
