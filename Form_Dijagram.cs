using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

using System.Web.UI.DataVisualization.Charting;

namespace TETOAplikacija
{
    public partial class Form_Dijagram : Form
    {
        public Form_Dijagram()
        {
            InitializeComponent();
        }

        double[,] mat;

        public double[,] rezultatiOptimizacije(double[,] rezultatiOptimizacije)
        {
            mat = rezultatiOptimizacije;
            crtanjeGrafa();
            return mat;
        }


        public void crtanjeGrafa()
        {


            Form_Glavna formaUlaz = new Form_Glavna();
            double pocetnaNapunjenostAkululatora=Convert.ToDouble(formaUlaz.txtNapunjenostSpremnika.Text);


            for (int i = 0; i < 24; i++)
            {

                if (i == 0)
                {
                    chartGraphic.Series["Akumulator"].Points.AddXY(i, pocetnaNapunjenostAkululatora);
                }
               

                chartGraphic.Series["Akumulator"].Points.AddXY(i+1, mat[i,55-1]);
                chartGraphic.Series["El. energija - K blok"].Points.AddXY(i + 1, mat[i, 12 - 1] + mat[i, 16 - 1] + mat[i, 32 - 1]);
                chartGraphic.Series["El. energija - L blok"].Points.AddXY(i + 1, mat[i, 21 - 1] + mat[i, 31 - 1]);
                chartGraphic.Series["El. energija - C blok"].Points.AddXY(i + 1, mat[i, 66 - 1]);


                chartGraphic2.Series["Toplina - K blok"].Points.AddXY(i + 1, mat[i, 13 - 1] + mat[i, 17 - 1] + mat[i, 29 - 1]);
                chartGraphic2.Series["Toplina - L blok"].Points.AddXY(i + 1, mat[i, 22 - 1] + mat[i, 44 - 1]);
                chartGraphic2.Series["Toplina - C blok"].Points.AddXY(i + 1, mat[i, 63 - 1]);
                chartGraphic2.Series["Toplina - VK"].Points.AddXY(i + 1, mat[i, 36 - 1]);
                chartGraphic2.Series["Toplina - C6"].Points.AddXY(i + 1, mat[i, 45 - 1]);


                chartGraphic.ChartAreas[0].AxisX.RoundAxisValues();

                chartGraphic.ChartAreas[0].AxisX.Minimum = 0;
                chartGraphic.ChartAreas[0].AxisX.Maximum = 24;
                chartGraphic.ChartAreas[0].AxisX.Interval = 1;


                chartGraphic2.ChartAreas[0].AxisX.Minimum = 0;
                chartGraphic2.ChartAreas[0].AxisX.Maximum = 24;
                chartGraphic2.ChartAreas[0].AxisX.Interval = 1;

                chartGraphic.ChartAreas[0].AxisX.Title = "Sat";
                chartGraphic.ChartAreas[0].AxisY.Title = "MWh";

                chartGraphic2.ChartAreas[0].AxisX.Title = "Sat";
                chartGraphic2.ChartAreas[0].AxisY.Title = "MWh";


            }
        }

        private void button1_Click(object sender, EventArgs e)
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













    }
}
