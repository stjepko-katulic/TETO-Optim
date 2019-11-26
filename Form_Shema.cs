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
    public partial class Form_Shema : Form
    {
        public Form_Shema()
        {
            InitializeComponent();
        }

        // deklariranje globalne varijable za ovu windows formu
        double[,] mat;



        public double[,] rezultatiOptimizacije(double[,] rezultatiOptimizacije)
        {
            mat = rezultatiOptimizacije;
            return mat;        
        }


        public void ispisSatnihRezultata(int sat)
        {

            sat--;

            txt_mgK1.Text = Convert.ToString(Math.Round(mat[sat, 8-1],2));
            txt_elOptK1.Text = Convert.ToString(Math.Round(mat[sat, 2 - 1], 2));
            txt_PptK1.Text = Convert.ToString(Math.Round(mat[sat, 12 - 1], 2));
            txt_QZVK1.Text = Convert.ToString(Math.Round(mat[sat, 13 - 1], 2));
            txt_D10K1.Text = Convert.ToString(Math.Round(mat[sat, 10 - 1], 2));
            txt_Dw5.Text = Convert.ToString(Math.Round(mat[sat, 27 - 1], 2));
            txt_D91K1.Text = Convert.ToString(Math.Round(mat[sat, 11 - 1], 2));
            txt_Dw6.Text = Convert.ToString(Math.Round(mat[sat, 26 - 1], 2));
            txt_DTAKin.Text = Convert.ToString(Math.Round(mat[sat, 11 - 1] + mat[sat, 15 - 1], 2));
            txt_DTAK1ex.Text = Convert.ToString(Math.Round(mat[sat, 3 - 1], 2));
            txt_DTAK2ex.Text = Convert.ToString(Math.Round(mat[sat, 4 - 1], 2));
            txt_PK.Text = Convert.ToString(Math.Round(mat[sat, 32 - 1], 2));
            txt_Dw7.Text = Convert.ToString(Math.Round(mat[sat, 25 - 1], 2));
            txt_D91K2.Text = Convert.ToString(Math.Round(mat[sat, 15 - 1], 2));
            txt_D10K2.Text = Convert.ToString(Math.Round(mat[sat, 14 - 1], 2));
            txt_QZVK2.Text = Convert.ToString(Math.Round(mat[sat, 17 - 1], 2));
            txt_DC4.Text = Convert.ToString(Math.Round(mat[sat, 41 - 1], 2));
            txt_QC4.Text = Convert.ToString(Math.Round(mat[sat, 29 - 1], 2));
            txt_PptK2.Text = Convert.ToString(Math.Round(mat[sat, 16 - 1], 2));
            txt_elOptK2.Text = Convert.ToString(Math.Round(mat[sat, 2 - 1], 2));
            txt_mgK2.Text = Convert.ToString(Math.Round(mat[sat, 9 - 1], 2));
            txt_QVK.Text = Convert.ToString(Math.Round(mat[sat, 36 - 1], 2));
            txt_mgVK.Text = Convert.ToString(Math.Round(mat[sat, 46 - 1], 2));
            txt_DL10.Text = Convert.ToString(Math.Round(mat[sat, 52 - 1], 2));
            txt_Dw8.Text = Convert.ToString(Math.Round(mat[sat, 24 - 1], 2));
            txt_Dw9.Text = Convert.ToString(Math.Round(mat[sat, 23 - 1], 2));

            txt_DC6.Text = Convert.ToString(Math.Round(mat[sat, 43 - 1], 2));
            txt_QC6.Text = Convert.ToString(Math.Round(mat[sat, 45 - 1], 2));

            txt_D10L.Text = Convert.ToString(Math.Round(mat[sat, 19 - 1], 2));
            txt_D95L.Text = Convert.ToString(Math.Round(mat[sat, 20 - 1], 2));
            txt_DTAL1ex.Text = Convert.ToString(Math.Round(mat[sat, 6 - 1], 2));
            txt_DTAL2ex.Text = Convert.ToString(Math.Round(mat[sat, 7 - 1], 2));
            txt_PL.Text = Convert.ToString(Math.Round(mat[sat, 31 - 1], 2));
            txt_QZVL.Text = Convert.ToString(Math.Round(mat[sat, 22 - 1], 2));
            txt_PptL.Text = Convert.ToString(Math.Round(mat[sat, 21 - 1], 2));
            txt_mgL.Text = Convert.ToString(Math.Round(mat[sat, 18 - 1], 2));
            txt_elOptL.Text = Convert.ToString(Math.Round(mat[sat, 5 - 1], 2));            
            txt_QC5.Text = Convert.ToString(Math.Round(mat[sat, 44 - 1], 2));
            txt_DC5.Text = Convert.ToString(Math.Round(mat[sat, 42 - 1], 2));

            txt_D140C.Text = Convert.ToString(Math.Round(mat[sat, 65 - 1], 2));
            txt_mgC.Text = Convert.ToString(Math.Round(mat[sat, 67 - 1], 2));
            txt_DTACin.Text = Convert.ToString(Math.Round(mat[sat, 64 - 1], 2));
            txt_optC.Text = Convert.ToString(Math.Round(mat[sat, 62 - 1], 2));
            txt_PelC.Text = Convert.ToString(Math.Round(mat[sat, 66 - 1], 2));
            txt_QC.Text = Convert.ToString(Math.Round(mat[sat, 63 - 1], 2));

            txt_IndPara.Text = Convert.ToString(Math.Round(mat[sat, 28 - 1], 2));


            txtToplinskiKonzum.Text = Convert.ToString(Math.Round(mat[sat, 36 - 1] + mat[sat, 56 - 1], 2));
            txtKonzumPara.Text = Convert.ToString(Math.Round(mat[sat, 28 - 1], 2));
            txtUPEL.Text = Convert.ToString(Math.Round(mat[sat, 79 - 1]/100, 2));
            txtUPEK.Text = Convert.ToString(Math.Round(mat[sat, 89 - 1] / 100, 2));
            txtProfit.Text = Convert.ToString(Math.Round(mat[sat, 1 - 1], 2));

        }




        private void satToolStripMenuItem_Click(object sender, EventArgs e)
        {
            ispisSatnihRezultata(1);
            txtSat.Text = "1. sat";
            
        }

        private void satToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            ispisSatnihRezultata(2);
            txtSat.Text = "2. sat";
        }

        private void satToolStripMenuItem2_Click(object sender, EventArgs e)
        {
            ispisSatnihRezultata(3);
            txtSat.Text = "3. sat";
        }

        private void satToolStripMenuItem3_Click(object sender, EventArgs e)
        {
            ispisSatnihRezultata(4);
            txtSat.Text = "4. sat";
        }

        private void satToolStripMenuItem4_Click(object sender, EventArgs e)
        {
            ispisSatnihRezultata(5);
            txtSat.Text = "5. sat";
        }

        private void satToolStripMenuItem5_Click(object sender, EventArgs e)
        {
            ispisSatnihRezultata(6);
            txtSat.Text = "6. sat";
        }

        private void satToolStripMenuItem6_Click(object sender, EventArgs e)
        {
            ispisSatnihRezultata(7);
            txtSat.Text = "7. sat";
        }

        private void satToolStripMenuItem7_Click(object sender, EventArgs e)
        {
            ispisSatnihRezultata(8);
            txtSat.Text = "8. sat";
        }

        private void satToolStripMenuItem8_Click(object sender, EventArgs e)
        {
            ispisSatnihRezultata(9);
            txtSat.Text = "9. sat";
        }

        private void satToolStripMenuItem9_Click(object sender, EventArgs e)
        {
            ispisSatnihRezultata(10);
            txtSat.Text = "10. sat";
        }

        private void satToolStripMenuItem10_Click(object sender, EventArgs e)
        {
            ispisSatnihRezultata(11);
            txtSat.Text = "11. sat";
        }

        private void satToolStripMenuItem11_Click(object sender, EventArgs e)
        {
            ispisSatnihRezultata(12);
            txtSat.Text = "12. sat";
        }

        private void satToolStripMenuItem12_Click(object sender, EventArgs e)
        {
            ispisSatnihRezultata(13);
            txtSat.Text = "13. sat";
        }

        private void satToolStripMenuItem13_Click(object sender, EventArgs e)
        {
            ispisSatnihRezultata(14);
            txtSat.Text = "14. sat";
        }

        private void satToolStripMenuItem14_Click(object sender, EventArgs e)
        {
            ispisSatnihRezultata(15);
            txtSat.Text = "15. sat";
        }

        private void satToolStripMenuItem15_Click(object sender, EventArgs e)
        {
            ispisSatnihRezultata(16);
            txtSat.Text = "16. sat";
        }

        private void satToolStripMenuItem16_Click(object sender, EventArgs e)
        {
            ispisSatnihRezultata(17);
            txtSat.Text = "17. sat";
        }

        private void satToolStripMenuItem17_Click(object sender, EventArgs e)
        {
            ispisSatnihRezultata(18);
            txtSat.Text = "18. sat";
        }

        private void satToolStripMenuItem18_Click(object sender, EventArgs e)
        {
            ispisSatnihRezultata(19);
            txtSat.Text = "19. sat";
        }

        private void satToolStripMenuItem19_Click(object sender, EventArgs e)
        {
            ispisSatnihRezultata(20);
            txtSat.Text = "20. sat";
        }

        private void satToolStripMenuItem20_Click(object sender, EventArgs e)
        {
            ispisSatnihRezultata(21);
            txtSat.Text = "21. sat";
        }

        private void satToolStripMenuItem21_Click(object sender, EventArgs e)
        {
            ispisSatnihRezultata(22);
            txtSat.Text = "22. sat";
        }

        private void satToolStripMenuItem22_Click(object sender, EventArgs e)
        {
            ispisSatnihRezultata(23);
            txtSat.Text = "23. sat";
        }

        private void satToolStripMenuItem23_Click(object sender, EventArgs e)
        {
            ispisSatnihRezultata(24);
            txtSat.Text = "24. sat";
        }




        // ova metoda sprema screennshot aktivnog prozora kao jpeg sliku
        private void shemaToolStripMenuItem_Click(object sender, EventArgs e)
        {
            SaveShema();
        }


        public void SaveShema()
        {
            SaveFileDialog dialog = new SaveFileDialog();
            string filePath;

            if (dialog.ShowDialog() == DialogResult.OK)
            {
                filePath = dialog.FileName;
                Bitmap myBitmap = new Bitmap(this.Width, this.Height);
                DrawToBitmap(myBitmap, new Rectangle(0, 0, myBitmap.Width, myBitmap.Height));
                myBitmap.Save(filePath + ".jpeg", System.Drawing.Imaging.ImageFormat.Jpeg);
            }
        }



        // otvara novi prozor (forma) s grafom
        private void dijagramToolStripMenuItem_Click(object sender, EventArgs e)
        {

            Form_Dijagram formaGraf = new Form_Dijagram();
            formaGraf.rezultatiOptimizacije(mat);
            formaGraf.Show();
            
        }

        private void zatvaranjeForme(object sender, FormClosingEventArgs e)
        {
            this.Hide();
            e.Cancel = true;
        }

        private void btnIconSaveShema_Click(object sender, EventArgs e)
        {
            SaveShema();
        }

        private void btnIconDiagram_Click(object sender, EventArgs e)
        {
            Form_Dijagram formaGraf = new Form_Dijagram();
            formaGraf.rezultatiOptimizacije(mat);
            formaGraf.Show();
        }





























    }
}
