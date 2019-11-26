namespace TETOAplikacija
{
    partial class Form_Dijagram
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            System.Windows.Forms.DataVisualization.Charting.ChartArea chartArea1 = new System.Windows.Forms.DataVisualization.Charting.ChartArea();
            System.Windows.Forms.DataVisualization.Charting.Legend legend1 = new System.Windows.Forms.DataVisualization.Charting.Legend();
            System.Windows.Forms.DataVisualization.Charting.Series series1 = new System.Windows.Forms.DataVisualization.Charting.Series();
            System.Windows.Forms.DataVisualization.Charting.Series series2 = new System.Windows.Forms.DataVisualization.Charting.Series();
            System.Windows.Forms.DataVisualization.Charting.Series series3 = new System.Windows.Forms.DataVisualization.Charting.Series();
            System.Windows.Forms.DataVisualization.Charting.Series series4 = new System.Windows.Forms.DataVisualization.Charting.Series();
            System.Windows.Forms.DataVisualization.Charting.ChartArea chartArea2 = new System.Windows.Forms.DataVisualization.Charting.ChartArea();
            System.Windows.Forms.DataVisualization.Charting.Legend legend2 = new System.Windows.Forms.DataVisualization.Charting.Legend();
            System.Windows.Forms.DataVisualization.Charting.Series series5 = new System.Windows.Forms.DataVisualization.Charting.Series();
            System.Windows.Forms.DataVisualization.Charting.Series series6 = new System.Windows.Forms.DataVisualization.Charting.Series();
            System.Windows.Forms.DataVisualization.Charting.Series series7 = new System.Windows.Forms.DataVisualization.Charting.Series();
            System.Windows.Forms.DataVisualization.Charting.Series series8 = new System.Windows.Forms.DataVisualization.Charting.Series();
            System.Windows.Forms.DataVisualization.Charting.Series series9 = new System.Windows.Forms.DataVisualization.Charting.Series();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form_Dijagram));
            this.chartGraphic = new System.Windows.Forms.DataVisualization.Charting.Chart();
            this.chartGraphic2 = new System.Windows.Forms.DataVisualization.Charting.Chart();
            this.button1 = new System.Windows.Forms.Button();
            this.toolTip_Ispisi = new System.Windows.Forms.ToolTip(this.components);
            this.panel1 = new System.Windows.Forms.Panel();
            ((System.ComponentModel.ISupportInitialize)(this.chartGraphic)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.chartGraphic2)).BeginInit();
            this.panel1.SuspendLayout();
            this.SuspendLayout();
            // 
            // chartGraphic
            // 
            this.chartGraphic.BorderlineColor = System.Drawing.Color.Gray;
            chartArea1.Name = "ChartArea1";
            this.chartGraphic.ChartAreas.Add(chartArea1);
            legend1.Name = "Legend1";
            this.chartGraphic.Legends.Add(legend1);
            this.chartGraphic.Location = new System.Drawing.Point(11, 54);
            this.chartGraphic.Name = "chartGraphic";
            series1.BorderWidth = 2;
            series1.ChartArea = "ChartArea1";
            series1.ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Line;
            series1.Color = System.Drawing.Color.Blue;
            series1.Legend = "Legend1";
            series1.Name = "Akumulator";
            series2.BorderWidth = 2;
            series2.ChartArea = "ChartArea1";
            series2.ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Line;
            series2.Color = System.Drawing.Color.Red;
            series2.LabelForeColor = System.Drawing.Color.Blue;
            series2.Legend = "Legend1";
            series2.Name = "El. energija - K blok";
            series3.BorderWidth = 2;
            series3.ChartArea = "ChartArea1";
            series3.ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Line;
            series3.Color = System.Drawing.Color.Lime;
            series3.LabelForeColor = System.Drawing.Color.Red;
            series3.Legend = "Legend1";
            series3.Name = "El. energija - L blok";
            series4.BorderWidth = 2;
            series4.ChartArea = "ChartArea1";
            series4.ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Line;
            series4.Color = System.Drawing.Color.Fuchsia;
            series4.LabelForeColor = System.Drawing.Color.Fuchsia;
            series4.Legend = "Legend1";
            series4.Name = "El. energija - C blok";
            this.chartGraphic.Series.Add(series1);
            this.chartGraphic.Series.Add(series2);
            this.chartGraphic.Series.Add(series3);
            this.chartGraphic.Series.Add(series4);
            this.chartGraphic.Size = new System.Drawing.Size(699, 468);
            this.chartGraphic.TabIndex = 0;
            this.chartGraphic.Text = "chart1";
            // 
            // chartGraphic2
            // 
            chartArea2.Name = "ChartArea1";
            this.chartGraphic2.ChartAreas.Add(chartArea2);
            legend2.Name = "Legend1";
            this.chartGraphic2.Legends.Add(legend2);
            this.chartGraphic2.Location = new System.Drawing.Point(734, 54);
            this.chartGraphic2.Name = "chartGraphic2";
            series5.BorderWidth = 2;
            series5.ChartArea = "ChartArea1";
            series5.ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Line;
            series5.Color = System.Drawing.Color.Blue;
            series5.Legend = "Legend1";
            series5.Name = "Toplina - K blok";
            series6.BorderWidth = 2;
            series6.ChartArea = "ChartArea1";
            series6.ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Line;
            series6.Color = System.Drawing.Color.Red;
            series6.Legend = "Legend1";
            series6.Name = "Toplina - L blok";
            series7.BorderWidth = 2;
            series7.ChartArea = "ChartArea1";
            series7.ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Line;
            series7.Color = System.Drawing.Color.Lime;
            series7.Legend = "Legend1";
            series7.Name = "Toplina - C blok";
            series8.BorderWidth = 2;
            series8.ChartArea = "ChartArea1";
            series8.ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Line;
            series8.Color = System.Drawing.Color.Fuchsia;
            series8.Legend = "Legend1";
            series8.Name = "Toplina - VK";
            series9.BorderWidth = 2;
            series9.ChartArea = "ChartArea1";
            series9.ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Line;
            series9.Color = System.Drawing.Color.Black;
            series9.Legend = "Legend1";
            series9.Name = "Toplina - C6";
            this.chartGraphic2.Series.Add(series5);
            this.chartGraphic2.Series.Add(series6);
            this.chartGraphic2.Series.Add(series7);
            this.chartGraphic2.Series.Add(series8);
            this.chartGraphic2.Series.Add(series9);
            this.chartGraphic2.Size = new System.Drawing.Size(672, 468);
            this.chartGraphic2.TabIndex = 1;
            this.chartGraphic2.Text = "chart1";
            // 
            // button1
            // 
            this.button1.BackColor = System.Drawing.SystemColors.ButtonHighlight;
            this.button1.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("button1.BackgroundImage")));
            this.button1.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.button1.Location = new System.Drawing.Point(6, 2);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(34, 34);
            this.button1.TabIndex = 6;
            this.toolTip_Ispisi.SetToolTip(this.button1, "Ispiši dijagram u jpg formatu.");
            this.button1.UseVisualStyleBackColor = false;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // panel1
            // 
            this.panel1.BackColor = System.Drawing.Color.Beige;
            this.panel1.Controls.Add(this.button1);
            this.panel1.Location = new System.Drawing.Point(0, 2);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(1533, 38);
            this.panel1.TabIndex = 7;
            // 
            // Form3
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.WhiteSmoke;
            this.ClientSize = new System.Drawing.Size(1424, 551);
            this.Controls.Add(this.panel1);
            this.Controls.Add(this.chartGraphic2);
            this.Controls.Add(this.chartGraphic);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximumSize = new System.Drawing.Size(1440, 590);
            this.MinimumSize = new System.Drawing.Size(1440, 590);
            this.Name = "Form3";
            this.Text = "Dijagram";
            ((System.ComponentModel.ISupportInitialize)(this.chartGraphic)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.chartGraphic2)).EndInit();
            this.panel1.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.DataVisualization.Charting.Chart chartGraphic;
        private System.Windows.Forms.DataVisualization.Charting.Chart chartGraphic2;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.ToolTip toolTip_Ispisi;
        private System.Windows.Forms.Panel panel1;
    }
}