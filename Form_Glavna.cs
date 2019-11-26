using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.Threading.Tasks;
using System.Threading;
using MathWorks.MATLAB.NET.Arrays;
using Microsoft.Office.Interop.Excel;
using optimizacijaTETO;


namespace TETOAplikacija
{
    public partial class Form_Glavna : Form
    {
        double[,] mat = new double[24, 113];
        MWArray MatricaRezultati;
        MWArray ukljucenostBlokaL;
        MWArray ukljucenostBlokaK;
        MWArray ukljucenostBlokaC;
        MWArray CijenaElEnergijeNT;
        MWArray CijenaElEnergijeVT;
        MWArray CijenaPlin;
        MWArray CijenaMazut;
        MWArray CijenaToplina;
        MWArray CijenaPara;
        MWArray pocNapunjenostSpremnika;
        MWArray brojTurbinaK;
        MWArray QVKGranica;
        MWArray UPEL;
        MWArray EtaL;
        Form_Shema formShema;
        Form_Opcije formaOpcije;



        public Form_Glavna()
        {
            InitializeComponent();
            formShema = new Form_Shema();
            formaOpcije = new Form_Opcije();
        }




        //----------------------------------------------------------------------------------------------------------------------       



        private async void btnOptimiraj_Click(object sender, EventArgs e)
        {
            this.Text = "TETO Zagreb    ------->    Simulacija optimizacija u tijeku ...";
            this.btnOptimiraj.Enabled = false;

            // kreiranje taska za simulaciju optimizacije
            Task myTask = new Task(new System.Action(() => SimulacijaOptimizacije()));
            myTask.Start();
            await myTask;

            // za stvarnu optimizaciju kreiran bi bio sljedeci task
            ////Task myTask = new Task(new System.Action(() => Optim()));
            ////myTask.Start();
            ////await myTask;

            this.Text = "TETO Zagreb";
            this.btnOptimiraj.Enabled = true;

            MessageBox.Show("Optimizacija je završila!", "Obavijest");
        }


        //----------------------------------------------------------------------------------------------------------------------   

        public void SimulacijaOptimizacije()
        {
            int trajanjeSimulacije = 10;

            BeginInvoke(new System.Action(() =>
            {
                this.Text = "TETO Zagreb    ------->    Simulacija optimizacija u tijeku ... " + trajanjeSimulacije;
            }));


            // ovdje simuliramo vrijeme potrebno za optimizaciju (npr. 10 sec)
            for (int i = trajanjeSimulacije; i > 0; i--)
            {
                Thread.Sleep(1000);
                BeginInvoke(new System.Action(() =>
                {
                    this.Text = "TETO Zagreb    ------->    Simulacija optimizacija u tijeku ... " + i;
                }));
            }
        }


        //----------------------------------------------------------------------------------------------------------------------   


        public void Optim()
        {
            try
            {
                // toplinski konzum u svakom satu (24 sata)
                MWArray TK1 = (MWArray)(Convert.ToDouble(txtTK1.Text));
                MWArray TK2 = (MWArray)(Convert.ToDouble(txtTK2.Text));
                MWArray TK3 = (MWArray)(Convert.ToDouble(txtTK3.Text));
                MWArray TK4 = (MWArray)(Convert.ToDouble(txtTK4.Text));
                MWArray TK5 = (MWArray)(Convert.ToDouble(txtTK5.Text));
                MWArray TK6 = (MWArray)(Convert.ToDouble(txtTK6.Text));
                MWArray TK7 = (MWArray)(Convert.ToDouble(txtTK7.Text));
                MWArray TK8 = (MWArray)(Convert.ToDouble(txtTK8.Text));
                MWArray TK9 = (MWArray)(Convert.ToDouble(txtTK9.Text));
                MWArray TK10 = (MWArray)(Convert.ToDouble(txtTK10.Text));
                MWArray TK11 = (MWArray)(Convert.ToDouble(txtTK11.Text));
                MWArray TK12 = (MWArray)(Convert.ToDouble(txtTK12.Text));
                MWArray TK13 = (MWArray)(Convert.ToDouble(txtTK13.Text));
                MWArray TK14 = (MWArray)(Convert.ToDouble(txtTK14.Text));
                MWArray TK15 = (MWArray)(Convert.ToDouble(txtTK15.Text));
                MWArray TK16 = (MWArray)(Convert.ToDouble(txtTK16.Text));
                MWArray TK17 = (MWArray)(Convert.ToDouble(txtTK17.Text));
                MWArray TK18 = (MWArray)(Convert.ToDouble(txtTK18.Text));
                MWArray TK19 = (MWArray)(Convert.ToDouble(txtTK19.Text));
                MWArray TK20 = (MWArray)(Convert.ToDouble(txtTK20.Text));
                MWArray TK21 = (MWArray)(Convert.ToDouble(txtTK21.Text));
                MWArray TK22 = (MWArray)(Convert.ToDouble(txtTK22.Text));
                MWArray TK23 = (MWArray)(Convert.ToDouble(txtTK23.Text));
                MWArray TK24 = (MWArray)(Convert.ToDouble(txtTK24.Text));

                // konzum ind. pare u svakom satu (24 sata)
                MWArray KP1 = (MWArray)(Convert.ToDouble(txtKP1.Text));
                MWArray KP2 = (MWArray)(Convert.ToDouble(txtKP2.Text));
                MWArray KP3 = (MWArray)(Convert.ToDouble(txtKP3.Text));
                MWArray KP4 = (MWArray)(Convert.ToDouble(txtKP4.Text));
                MWArray KP5 = (MWArray)(Convert.ToDouble(txtKP5.Text));
                MWArray KP6 = (MWArray)(Convert.ToDouble(txtKP6.Text));
                MWArray KP7 = (MWArray)(Convert.ToDouble(txtKP7.Text));
                MWArray KP8 = (MWArray)(Convert.ToDouble(txtKP8.Text));
                MWArray KP9 = (MWArray)(Convert.ToDouble(txtKP9.Text));
                MWArray KP10 = (MWArray)(Convert.ToDouble(txtKP10.Text));
                MWArray KP11 = (MWArray)(Convert.ToDouble(txtKP11.Text));
                MWArray KP12 = (MWArray)(Convert.ToDouble(txtKP12.Text));
                MWArray KP13 = (MWArray)(Convert.ToDouble(txtKP13.Text));
                MWArray KP14 = (MWArray)(Convert.ToDouble(txtKP14.Text));
                MWArray KP15 = (MWArray)(Convert.ToDouble(txtKP15.Text));
                MWArray KP16 = (MWArray)(Convert.ToDouble(txtKP16.Text));
                MWArray KP17 = (MWArray)(Convert.ToDouble(txtKP17.Text));
                MWArray KP18 = (MWArray)(Convert.ToDouble(txtKP18.Text));
                MWArray KP19 = (MWArray)(Convert.ToDouble(txtKP19.Text));
                MWArray KP20 = (MWArray)(Convert.ToDouble(txtKP20.Text));
                MWArray KP21 = (MWArray)(Convert.ToDouble(txtKP21.Text));
                MWArray KP22 = (MWArray)(Convert.ToDouble(txtKP22.Text));
                MWArray KP23 = (MWArray)(Convert.ToDouble(txtKP23.Text));
                MWArray KP24 = (MWArray)(Convert.ToDouble(txtKP24.Text));

                // ostali parametri
                MWArray CijenaElEnergijeNT = (MWArray)(Convert.ToDouble(txtCijenaElEnergijeNT.Text));
                MWArray CijenaElEnergijeVT = (MWArray)(Convert.ToDouble(txtCijenaElEnergijeVT.Text));
                MWArray CijenaPlin = (MWArray)(Convert.ToDouble(txtCijenaPlin.Text));
                MWArray CijenaMazut = (MWArray)(Convert.ToDouble(txtCijenaMazut.Text));
                MWArray CijenaToplina = (MWArray)(Convert.ToDouble(txtCijenaToplina.Text));
                MWArray CijenaPara = (MWArray)(Convert.ToDouble(txtCijenaPara.Text));
                MWArray pocNapunjenostSpremnika = (MWArray)(Convert.ToDouble(txtNapunjenostSpremnika.Text));
                MWArray brojTurbinaK = (MWArray)(Convert.ToDouble(txtBrojTurbina.Text));
                MWArray QVKGranica = (MWArray)(Convert.ToDouble(txtGranicaPaljenjaVK.Text));
                MWArray UPEL = (MWArray)(Convert.ToDouble(txtUPEL.Text));
                MWArray EtaL = (MWArray)(Convert.ToDouble(txtEtaL.Text));
                MWArray TlakIndPara = (MWArray)(Convert.ToDouble(formaOpcije.txtTlakIndPara.Text));
                MWArray TlakPareC = (MWArray)(Convert.ToDouble(formaOpcije.txtTlakPareC.Text));
                MWArray TlakPareL = (MWArray)(Convert.ToDouble(formaOpcije.txtTlakPareL.Text));
                MWArray TlakPareK = (MWArray)(Convert.ToDouble(formaOpcije.txtTlakPareK.Text));
                MWArray Para2Oduz = (MWArray)(Convert.ToDouble(formaOpcije.txtPara2oduz.Text));

                MWArray TempIndPare = (MWArray)(Convert.ToDouble(formaOpcije.txtTempIndPare.Text));
                MWArray TempVTParaK = (MWArray)(Convert.ToDouble(formaOpcije.txtTempVTParaK.Text));
                MWArray TempNTParaK = (MWArray)(Convert.ToDouble(formaOpcije.txtTempNTParaK.Text));
                MWArray TempVTParaL = (MWArray)(Convert.ToDouble(formaOpcije.txtTempVTParaL.Text));
                MWArray TempNTParaL = (MWArray)(Convert.ToDouble(formaOpcije.txtTempNTParaL.Text));
                MWArray TempVTParaC = (MWArray)(Convert.ToDouble(formaOpcije.txtTempVTParaC.Text));

                MWArray Algoritam = (MWArray)(Convert.ToDouble(formaOpcije.comboBox1.SelectedIndex));
                MWArray BrojIteracija = (MWArray)(Convert.ToDouble(formaOpcije.txtBrojIteracija.Text));
                MWArray TolCon = (MWArray)(Convert.ToDouble(formaOpcije.txtTolCon.Text));
                MWArray TolFun = (MWArray)(Convert.ToDouble(formaOpcije.txtTolFun.Text));
                MWArray TolX = (MWArray)(Convert.ToDouble(formaOpcije.txtTolX.Text));



                // provjerava se da li je blok L u pogonu
                bool checkBlokL = cbxLBlok.Checked;

                if (checkBlokL == true)
                {
                    ukljucenostBlokaL = (MWArray)(Convert.ToDouble(1));
                }
                else
                {
                    ukljucenostBlokaL = (MWArray)(Convert.ToDouble(0));
                }

                // provjerava se da li je blok K u pogonu
                bool checkBlokK = cbxKBlok.Checked;

                if (checkBlokK == true)
                {
                    ukljucenostBlokaK = (MWArray)(Convert.ToDouble(1));
                }
                else
                {
                    ukljucenostBlokaK = (MWArray)(Convert.ToDouble(0));
                }

                // provjerava se da li je blok C u pogonu
                bool checkBlokC = cbxCBlok.Checked;

                if (checkBlokC == true)
                {
                    ukljucenostBlokaC = (MWArray)(Convert.ToDouble(1));
                }
                else
                {
                    ukljucenostBlokaC = (MWArray)(Convert.ToDouble(0));
                }

                // poziva se metoda za provjeru ispravnosti upisanih podataka
                int izlaz = provjeraIspravnostiUpisanihPodataka();

                if (izlaz == 0)
                {
                    return;
                }




                //POKRETANJE OPTIMIZACIJE

                TETO optimizacija = new TETO();

                try
                {
                    MatricaRezultati = optimizacija.optimizacijaTETO(TK1, TK2, TK3, TK4, TK5, TK6, TK7, TK8, TK9, TK10, TK11, TK12, TK13, TK14, TK15,
                    TK16, TK17, TK18, TK19, TK20, TK21, TK22, TK23, TK24, KP1, KP2, KP3, KP4, KP5, KP6, KP7, KP8, KP9, KP10, KP11, KP12, KP13, KP14, KP15,
                    KP16, KP17, KP18, KP19, KP20, KP21, KP22, KP23, KP24, CijenaElEnergijeNT, CijenaElEnergijeVT, CijenaPlin, CijenaMazut, CijenaToplina,
                    CijenaPara, pocNapunjenostSpremnika, ukljucenostBlokaK, ukljucenostBlokaL, ukljucenostBlokaC, brojTurbinaK, QVKGranica, UPEL, EtaL,
                    TlakIndPara, TlakPareC, TlakPareL, TlakPareK, Para2Oduz, TempIndPare, TempVTParaK, TempNTParaK, TempNTParaL, TempVTParaL, TempVTParaC,
                    Algoritam, TolCon, TolFun, TolX, BrojIteracija);
                }
                catch
                {
                    MessageBox.Show("Program ne može naći fizikalno rješenje!");
                    this.Text = "TETO Zagreb";
                    return;
                }


                double[] vec = (double[])((MWNumericArray)MatricaRezultati).ToVector(MWArrayComponent.Real);

                // pretvaranje vektora rezultata u matricu
                int s = -1;

                for (int j = 0; j < 113; j++)
                {
                    for (int i = 0; i < 24; i++)
                    {
                        s++;
                        mat[i, j] = vec[s];
                    }
                }

                this.Text = "TETO Zagreb";
            }
            catch
            {
                MessageBox.Show("Podaci nisu unešeni u pravilnom obliku!");
                this.Text = "TETO Zagreb";
                return;
            }
        }


        //---------------------------------------------------------------------------------------------------------------------- 



        private void ispisUExcelToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            // ispisuju se rezultati optimizacije u Excel
            IspisExcel.IspisivanjeRezultataUExcel(mat);
        }


        //---------------------------------------------------------------------------------------------------------------------- 


        private int provjeraIspravnostiUpisanihPodataka() // provjerava da li su svi podaci relevantni za optimizaciju unešeni u odgovarajućem formatu
        {

            double CijenaElEnergijeNT_d = Convert.ToDouble(txtCijenaElEnergijeNT.Text);
            double CijenaElEnergijeVT_d = Convert.ToDouble(txtCijenaElEnergijeVT.Text);
            double CijenaPlin_d = Convert.ToDouble(txtCijenaPlin.Text);
            double CijenaMazut_d = Convert.ToDouble(txtCijenaMazut.Text);
            double CijenaToplina_d = Convert.ToDouble(txtCijenaToplina.Text);
            double CijenaPara_d = Convert.ToDouble(txtCijenaPara.Text);

            double brojTurbinaK_d = Convert.ToDouble(txtBrojTurbina.Text);
            double QVKGranica_d = Convert.ToDouble(txtGranicaPaljenjaVK.Text);


            double TlakIndPara_d = (Convert.ToDouble(formaOpcije.txtTlakIndPara.Text));
            double TlakPareC_d = (Convert.ToDouble(formaOpcije.txtTlakPareC.Text));
            double TlakPareL_d = (Convert.ToDouble(formaOpcije.txtTlakPareL.Text));
            double TlakPareK_d = (Convert.ToDouble(formaOpcije.txtTlakPareK.Text));
            double Para2Oduz_d = (Convert.ToDouble(formaOpcije.txtPara2oduz.Text));

            double TempIndPare_d = (Convert.ToDouble(formaOpcije.txtTempIndPare.Text));
            double TempVTParaK_d = (Convert.ToDouble(formaOpcije.txtTempVTParaK.Text));
            double TempNTParaK_d = (Convert.ToDouble(formaOpcije.txtTempNTParaK.Text));
            double TempVTParaL_d = (Convert.ToDouble(formaOpcije.txtTempVTParaL.Text));
            double TempNTParaL_d = (Convert.ToDouble(formaOpcije.txtTempNTParaL.Text));
            double TempVTParaC_d = (Convert.ToDouble(formaOpcije.txtTempVTParaC.Text));

            double BrojIteracija_d = (Convert.ToDouble(formaOpcije.txtBrojIteracija.Text));
            double TolCon_d = (Convert.ToDouble(formaOpcije.txtTolCon.Text));
            double TolFun_d = (Convert.ToDouble(formaOpcije.txtTolFun.Text));
            double TolX_d = (Convert.ToDouble(formaOpcije.txtTolX.Text));


            int izlaz;

            // provijera cijene elektricne energije
            if ((CijenaElEnergijeNT_d < 0) || (CijenaElEnergijeVT_d < 0))
            {
                MessageBox.Show("Cijena električne energije ne smije biti negativna!");
                izlaz = 0;
                this.Text = "TETO Zagreb";
                return izlaz;
            }


            // provijera cijene goriva
            if ((CijenaPlin_d < 0) || (CijenaMazut_d < 0))
            {
                MessageBox.Show("Cijena goriva ne smije biti negativna!");
                izlaz = 0;
                this.Text = "TETO Zagreb";
                return izlaz;
            }


            // provijera cijene topline i pare
            if ((CijenaToplina_d < 0) || (CijenaPara_d < 0))
            {
                MessageBox.Show("Cijena topline/pare ne smije biti negativna!");
                izlaz = 0;
                this.Text = "TETO Zagreb";
                return izlaz;
            }


            // provijera broja uklučenih turbina K bloka
            if (!(brojTurbinaK_d == 1) && !(brojTurbinaK_d == 2))
            {
                MessageBox.Show("Nepravilno unešen broj plinskih turnina u radu (K blok)!");
                izlaz = 0;
                this.Text = "TETO Zagreb";
                return izlaz;
            }


            if ((cbxKBlok.Checked == false) && (cbxLBlok.Checked == false) && (cbxCBlok.Checked == false))
            {
                MessageBox.Show("Minimalno jedno postrojenje treba raditi!");
                izlaz = 0;
                this.Text = "TETO Zagreb";
                return izlaz;
            }

            // provijera granice paljenja vršnog kotla
            if ((QVKGranica_d > 350) || (QVKGranica_d < 5))
            {
                MessageBox.Show("Granica paljenja vršnog/vrelovodnog kotla mora biti 5 MW<Qvk<350 MW!");
                izlaz = 0;
                this.Text = "TETO Zagreb";
                return izlaz;
            }


            if ((TlakIndPara_d > 30) || (TlakIndPara_d < 5))
            {
                formaOpcije.Show();
                MessageBox.Show("Opcije ---> Mora biti:  4 < Tlak ind. pare < 30 ");
                izlaz = 0;
                this.Text = "TETO Zagreb";
                return izlaz;
            }

            if ((TlakPareC_d > 300) || (TlakPareC_d < 60))
            {
                formaOpcije.Show();
                MessageBox.Show("Opcije ---> Mora biti:  60 < Tlak pare (C blok) < 300 ");
                izlaz = 0;
                this.Text = "TETO Zagreb";
                return izlaz;
            }

            if ((TlakPareL_d > 150) || (TlakPareL_d < 60))
            {
                formaOpcije.Show();
                MessageBox.Show("Opcije ---> Mora biti:  60 < Tlak pare (L blok) < 150 ");
                izlaz = 0;
                this.Text = "TETO Zagreb";
                return izlaz;
            }


            if ((TlakPareK_d > 150) || (TlakPareK_d < 60))
            {
                formaOpcije.Show();
                MessageBox.Show("Opcije ---> Mora biti:  60 < Tlak pare (K blok) < 150 ");
                izlaz = 0;
                this.Text = "TETO Zagreb";
                return izlaz;
            }



            if ((Para2Oduz_d > 10) || (Para2Oduz_d < 1))
            {
                formaOpcije.Show();
                MessageBox.Show("Opcije ---> Mora biti:  1 < Tlak pare 2. oduzimanje < 10 ");
                izlaz = 0;
                this.Text = "TETO Zagreb";
                return izlaz;
            }



            if ((TempIndPare_d > 300) || (TempIndPare_d < 120))
            {
                formaOpcije.Show();
                MessageBox.Show("Opcije ---> Mora biti:  120 < Temp. ind pare < 300 ");
                izlaz = 0;
                this.Text = "TETO Zagreb";
                return izlaz;
            }



            if ((TempVTParaK_d > 620) || (TempVTParaK_d < 400))
            {
                formaOpcije.Show();
                MessageBox.Show("Opcije ---> Mora biti:  400 < Temp. VT pare (K blok) < 620 ");
                izlaz = 0;
                this.Text = "TETO Zagreb";
                return izlaz;
            }



            if ((TempNTParaK_d > 350) || (TempNTParaK_d < 200))
            {
                formaOpcije.Show();
                MessageBox.Show("Opcije ---> Mora biti:  200 < Temp. NT pare (K blok) < 350 ");
                izlaz = 0;
                this.Text = "TETO Zagreb";
                return izlaz;
            }



            if ((TempVTParaL_d > 620) || (TempVTParaL_d < 400))
            {
                formaOpcije.Show();
                MessageBox.Show("Opcije ---> Mora biti:  400 < Temp. VT pare (L blok) < 620 ");
                izlaz = 0;
                this.Text = "TETO Zagreb";
                return izlaz;
            }



            if ((TempNTParaL_d > 350) || (TempNTParaL_d < 200))
            {
                formaOpcije.Show();
                MessageBox.Show("Opcije ---> Mora biti:  350 < Temp. NT pare (L blok) < 200 ");
                izlaz = 0;
                this.Text = "TETO Zagreb";
                return izlaz;
            }



            if ((TempVTParaC_d > 620) || (TempVTParaC_d < 400))
            {
                formaOpcije.Show();
                MessageBox.Show("Opcije ---> Mora biti:  400 < Temp. VT pare (C blok) < 620 ");
                izlaz = 0;
                this.Text = "TETO Zagreb";
                return izlaz;
            }



            if ((BrojIteracija_d > 30) || (BrojIteracija_d < 5))
            {
                formaOpcije.Show();
                MessageBox.Show("Opcije ---> Mora biti:  5 < Max. broj iteracija < 30 ");
                izlaz = 0;
                this.Text = "TETO Zagreb";
                return izlaz;
            }


            if ((TolCon_d > 0.001) || (TolCon_d < 0.00000001))
            {
                formaOpcije.Show();
                MessageBox.Show("Opcije ---> Mora biti:  0.001 < TolCon < 0.00000001 ");
                izlaz = 0;
                this.Text = "TETO Zagreb";
                return izlaz;
            }


            if ((TolFun_d > 0.001) || (TolFun_d < 0.00000001))
            {
                formaOpcije.Show();
                MessageBox.Show("Opcije ---> Mora biti:  0.001 < TolFun < 0.00000001 ");
                izlaz = 0;
                this.Text = "TETO Zagreb";
                return izlaz;
            }


            if ((TolX_d > 0.001) || (TolX_d < 0.00000001))
            {
                formaOpcije.Show();
                MessageBox.Show("Opcije ---> Mora biti:  0.001 < TolX < 0.00000001 ");
                izlaz = 0;
                this.Text = "TETO Zagreb";
                return izlaz;
            }


            if ((Convert.ToDouble(txtTK1.Text) < 0) ||
                (Convert.ToDouble(txtTK2.Text) < 0) ||
                (Convert.ToDouble(txtTK3.Text) < 0) ||
                (Convert.ToDouble(txtTK4.Text) < 0) ||
                (Convert.ToDouble(txtTK5.Text) < 0) ||
                (Convert.ToDouble(txtTK6.Text) < 0) ||
                (Convert.ToDouble(txtTK7.Text) < 0) ||
                (Convert.ToDouble(txtTK8.Text) < 0) ||
                (Convert.ToDouble(txtTK9.Text) < 0) ||
                (Convert.ToDouble(txtTK10.Text) < 0) ||
                (Convert.ToDouble(txtTK11.Text) < 0) ||
                (Convert.ToDouble(txtTK12.Text) < 0) ||
                (Convert.ToDouble(txtTK13.Text) < 0) ||
                (Convert.ToDouble(txtTK14.Text) < 0) ||
                (Convert.ToDouble(txtTK15.Text) < 0) ||
                (Convert.ToDouble(txtTK16.Text) < 0) ||
                (Convert.ToDouble(txtTK17.Text) < 0) ||
                (Convert.ToDouble(txtTK18.Text) < 0) ||
                (Convert.ToDouble(txtTK19.Text) < 0) ||
                (Convert.ToDouble(txtTK20.Text) < 0) ||
                (Convert.ToDouble(txtTK21.Text) < 0) ||
                (Convert.ToDouble(txtTK22.Text) < 0) ||
                (Convert.ToDouble(txtTK23.Text) < 0) ||
                (Convert.ToDouble(txtTK24.Text) < 0))
            {
                MessageBox.Show("Zadani toplinski konzumi moraju biti pozitivni!", "Upozorenje");
                izlaz = 0;
                this.Text = "TETO Zagreb";
                return izlaz;
            }


            // konzum ind. pare
            if ((Convert.ToDouble(txtKP1.Text) < 0) ||
            (Convert.ToDouble(txtKP2.Text) < 0) ||
            (Convert.ToDouble(txtKP3.Text) < 0) ||
            (Convert.ToDouble(txtKP4.Text) < 0) ||
            (Convert.ToDouble(txtKP5.Text) < 0) ||
            (Convert.ToDouble(txtKP6.Text) < 0) ||
            (Convert.ToDouble(txtKP7.Text) < 0) ||
            (Convert.ToDouble(txtKP8.Text) < 0) ||
            (Convert.ToDouble(txtKP9.Text) < 0) ||
            (Convert.ToDouble(txtKP10.Text) < 0) ||
            (Convert.ToDouble(txtKP11.Text) < 0) ||
            (Convert.ToDouble(txtKP12.Text) < 0) ||
            (Convert.ToDouble(txtKP13.Text) < 0) ||
            (Convert.ToDouble(txtKP14.Text) < 0) ||
            (Convert.ToDouble(txtKP15.Text) < 0) ||
            (Convert.ToDouble(txtKP16.Text) < 0) ||
            (Convert.ToDouble(txtKP17.Text) < 0) ||
            (Convert.ToDouble(txtKP18.Text) < 0) ||
            (Convert.ToDouble(txtKP19.Text) < 0) ||
            (Convert.ToDouble(txtKP20.Text) < 0) ||
            (Convert.ToDouble(txtKP21.Text) < 0) ||
            (Convert.ToDouble(txtKP22.Text) < 0) ||
            (Convert.ToDouble(txtKP23.Text) < 0) ||
            (Convert.ToDouble(txtKP24.Text) < 0))
            {
                MessageBox.Show("Zadani konzumi pare moraju biti pozitivni!");
                izlaz = 0;
                this.Text = "TETO Zagreb";
                return izlaz;
            }

            izlaz = 1;
            return izlaz;

        }

        //----------------------------------------------------------------------------------------------------------------------        


        private void unosKonzumaIzExcelaToolStripMenuItem_Click(object sender, EventArgs e) // podaci za toplinski konzum učitavaju se iz odabrane Excel datoteke
        {
            UnosKonzuma();
        }





        public void UnosKonzuma()
        {
            Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
            string filePath;
            Microsoft.Office.Interop.Excel.Workbook xlWorkbook;

            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                filePath = openFileDialog1.FileName;
            }
            else
            {
                return;
            }


            string fileExtension = Path.GetExtension(filePath);

            if ((fileExtension != ".xls") && (fileExtension != ".xlsm") && (fileExtension != ".xlsx"))
            {
                MessageBox.Show("Odabrana datoteka mora biti u Excel formatu!");
                return;
            }


            xlWorkbook = xlApp.Workbooks.Open(filePath);
            Microsoft.Office.Interop.Excel.Worksheet xlWorksheet = xlWorkbook.Sheets[1];
            Microsoft.Office.Interop.Excel.Range range;


            // unos podataka za toplinski konzum

            range = xlWorksheet.get_Range("A1");
            txtTK1.Text = range.Text;

            range = xlWorksheet.get_Range("A2");
            txtTK2.Text = range.Text;

            range = xlWorksheet.get_Range("A3");
            txtTK3.Text = range.Text;

            range = xlWorksheet.get_Range("A4");
            txtTK4.Text = range.Text;

            range = xlWorksheet.get_Range("A5");
            txtTK5.Text = range.Text;

            range = xlWorksheet.get_Range("A6");
            txtTK6.Text = range.Text;

            range = xlWorksheet.get_Range("A7");
            txtTK7.Text = range.Text;

            range = xlWorksheet.get_Range("A8");
            txtTK8.Text = range.Text;

            range = xlWorksheet.get_Range("A9");
            txtTK9.Text = range.Text;

            range = xlWorksheet.get_Range("A10");
            txtTK10.Text = range.Text;

            range = xlWorksheet.get_Range("A11");
            txtTK11.Text = range.Text;

            range = xlWorksheet.get_Range("A12");
            txtTK12.Text = range.Text;

            range = xlWorksheet.get_Range("A13");
            txtTK13.Text = range.Text;

            range = xlWorksheet.get_Range("A14");
            txtTK14.Text = range.Text;

            range = xlWorksheet.get_Range("A15");
            txtTK15.Text = range.Text;

            range = xlWorksheet.get_Range("A16");
            txtTK16.Text = range.Text;

            range = xlWorksheet.get_Range("A17");
            txtTK17.Text = range.Text;

            range = xlWorksheet.get_Range("A18");
            txtTK18.Text = range.Text;

            range = xlWorksheet.get_Range("A19");
            txtTK19.Text = range.Text;

            range = xlWorksheet.get_Range("A20");
            txtTK20.Text = range.Text;

            range = xlWorksheet.get_Range("A21");
            txtTK21.Text = range.Text;

            range = xlWorksheet.get_Range("A22");
            txtTK22.Text = range.Text;

            range = xlWorksheet.get_Range("A23");
            txtTK23.Text = range.Text;

            range = xlWorksheet.get_Range("A24");
            txtTK24.Text = range.Text;



            // unos podataka za konzum industrijske pare

            range = xlWorksheet.get_Range("B1");
            txtKP1.Text = range.Text;

            range = xlWorksheet.get_Range("B2");
            txtKP2.Text = range.Text;

            range = xlWorksheet.get_Range("B3");
            txtKP3.Text = range.Text;

            range = xlWorksheet.get_Range("B4");
            txtKP4.Text = range.Text;

            range = xlWorksheet.get_Range("B5");
            txtKP5.Text = range.Text;

            range = xlWorksheet.get_Range("B6");
            txtKP6.Text = range.Text;

            range = xlWorksheet.get_Range("B7");
            txtKP7.Text = range.Text;

            range = xlWorksheet.get_Range("B8");
            txtKP8.Text = range.Text;

            range = xlWorksheet.get_Range("B9");
            txtKP9.Text = range.Text;

            range = xlWorksheet.get_Range("B10");
            txtKP10.Text = range.Text;

            range = xlWorksheet.get_Range("B11");
            txtKP11.Text = range.Text;

            range = xlWorksheet.get_Range("B12");
            txtKP12.Text = range.Text;

            range = xlWorksheet.get_Range("B13");
            txtKP13.Text = range.Text;

            range = xlWorksheet.get_Range("B14");
            txtKP14.Text = range.Text;

            range = xlWorksheet.get_Range("B15");
            txtKP15.Text = range.Text;

            range = xlWorksheet.get_Range("B16");
            txtKP16.Text = range.Text;

            range = xlWorksheet.get_Range("B17");
            txtKP17.Text = range.Text;

            range = xlWorksheet.get_Range("B18");
            txtKP18.Text = range.Text;

            range = xlWorksheet.get_Range("B19");
            txtKP19.Text = range.Text;

            range = xlWorksheet.get_Range("B20");
            txtKP20.Text = range.Text;

            range = xlWorksheet.get_Range("B21");
            txtKP21.Text = range.Text;

            range = xlWorksheet.get_Range("B22");
            txtKP22.Text = range.Text;

            range = xlWorksheet.get_Range("B23");
            txtKP23.Text = range.Text;

            range = xlWorksheet.get_Range("B24");
            txtKP24.Text = range.Text;

            xlWorkbook.Close();
        }





        //----------------------------------------------------------------------------------------------------------------------  

        // prikazuje se novi prozor s prikazanom shemom postrojenja

        private void shemaPostrojenjaToolStripMenuItem_Click(object sender, EventArgs e)
        {
            formShema.rezultatiOptimizacije(mat);
            formShema.Show();
        }


        //----------------------------------------------------------------------------------------------------------------------  



        private void spremiRezultateToolStripMenuItem_Click(object sender, EventArgs e)
        {
            SaveData();
        }

        //----------------------------------------------------------------------------------------------------------------------  

        private void učitajRezultateToolStripMenuItem_Click(object sender, EventArgs e)
        {
            ReadData();
        }


        //----------------------------------------------------------------------------------------------------------------------  


        // spremaju se podaci u odabrani file (file ima ekstenziju *.data)
        public void SaveData()
        {
            string filePath;
            SaveFileDialog sfd = new SaveFileDialog();

            if (sfd.ShowDialog() == DialogResult.OK)
            {
                if (Path.GetExtension(sfd.FileName) == ".data")
                {
                    filePath = sfd.FileName;
                }
                else
                {
                    filePath = sfd.FileName + ".data";
                }
            }
            else
            {
                return;
            }


            StreamWriter sw = new StreamWriter(filePath);

            // spremaju se podaci varijable mat
            for (int i = 0; i < 24; i++)
            {
                for (int j = 0; j < 111; j++)
                {
                    sw.WriteLine(mat[i, j]);
                }
            }


            // spremaju se toplinski konzumi
            sw.WriteLine(txtTK1.Text);
            sw.WriteLine(txtTK2.Text);
            sw.WriteLine(txtTK3.Text);
            sw.WriteLine(txtTK4.Text);
            sw.WriteLine(txtTK5.Text);
            sw.WriteLine(txtTK6.Text);
            sw.WriteLine(txtTK7.Text);
            sw.WriteLine(txtTK8.Text);
            sw.WriteLine(txtTK9.Text);
            sw.WriteLine(txtTK10.Text);
            sw.WriteLine(txtTK11.Text);
            sw.WriteLine(txtTK12.Text);
            sw.WriteLine(txtTK13.Text);
            sw.WriteLine(txtTK14.Text);
            sw.WriteLine(txtTK15.Text);
            sw.WriteLine(txtTK16.Text);
            sw.WriteLine(txtTK17.Text);
            sw.WriteLine(txtTK18.Text);
            sw.WriteLine(txtTK19.Text);
            sw.WriteLine(txtTK20.Text);
            sw.WriteLine(txtTK21.Text);
            sw.WriteLine(txtTK22.Text);
            sw.WriteLine(txtTK23.Text);
            sw.WriteLine(txtTK24.Text);


            // spremaju se konzumi industrijske pare
            sw.WriteLine(txtKP1.Text);
            sw.WriteLine(txtKP2.Text);
            sw.WriteLine(txtKP3.Text);
            sw.WriteLine(txtKP4.Text);
            sw.WriteLine(txtKP5.Text);
            sw.WriteLine(txtKP6.Text);
            sw.WriteLine(txtKP7.Text);
            sw.WriteLine(txtKP8.Text);
            sw.WriteLine(txtKP9.Text);
            sw.WriteLine(txtKP10.Text);
            sw.WriteLine(txtKP11.Text);
            sw.WriteLine(txtKP12.Text);
            sw.WriteLine(txtKP13.Text);
            sw.WriteLine(txtKP14.Text);
            sw.WriteLine(txtKP15.Text);
            sw.WriteLine(txtKP16.Text);
            sw.WriteLine(txtKP17.Text);
            sw.WriteLine(txtKP18.Text);
            sw.WriteLine(txtKP19.Text);
            sw.WriteLine(txtKP20.Text);
            sw.WriteLine(txtKP21.Text);
            sw.WriteLine(txtKP22.Text);
            sw.WriteLine(txtKP23.Text);
            sw.WriteLine(txtKP24.Text);

            // spremaju se ostali podaci

            sw.WriteLine(txtCijenaElEnergijeNT.Text);
            sw.WriteLine(txtCijenaElEnergijeVT.Text);
            sw.WriteLine(txtCijenaPlin.Text);
            sw.WriteLine(txtCijenaMazut.Text);

            sw.WriteLine(txtCijenaToplina.Text);
            sw.WriteLine(txtCijenaPara.Text);

            sw.WriteLine(txtBrojTurbina.Text);
            sw.WriteLine(txtNapunjenostSpremnika.Text);



            if (cbxKBlok.Checked)
            {
                sw.WriteLine("true");
            }
            else
            {
                sw.WriteLine("false");
            }


            if (cbxLBlok.Checked)
            {
                sw.WriteLine("true");
            }
            else
            {
                sw.WriteLine("false");
            }


            if (cbxCBlok.Checked)
            {
                sw.WriteLine("true");
            }
            else
            {
                sw.WriteLine("false");
            }


            sw.Close();
            return;
        }

        //----------------------------------------------------------------------------------------------------------------------  

        // ucitava podatke iz odabranog fajla (fajl mora imati ekstenziju *.data)
        public void ReadData([System.Runtime.CompilerServices.CallerMemberName] string caller = null)
        {

            string filePath = Directory.GetCurrentDirectory() + "\\rezultati.data";

            if (!caller.Equals("SimulacijaOptimizacije"))
            {
                if (openFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    filePath = openFileDialog1.FileName;
                }
                else
                {
                    return;
                }

                string fileExtension = Path.GetExtension(filePath);

                if (fileExtension != ".data")
                {
                    MessageBox.Show("Odabrana datoteka mora biti u odgovarajućem formatu!");
                    return;
                }
            }



            StreamReader sr = new StreamReader(filePath);

            for (int i = 0; i < 24; i++)
            {
                for (int j = 0; j < 111; j++)
                {
                    mat[i, j] = Convert.ToDouble(sr.ReadLine());
                }
            }


            // učitavaju se toplinski konzumi
            txtTK1.Text = sr.ReadLine();
            txtTK2.Text = sr.ReadLine();
            txtTK3.Text = sr.ReadLine();
            txtTK4.Text = sr.ReadLine();
            txtTK5.Text = sr.ReadLine();
            txtTK6.Text = sr.ReadLine();
            txtTK7.Text = sr.ReadLine();
            txtTK8.Text = sr.ReadLine();
            txtTK9.Text = sr.ReadLine();
            txtTK10.Text = sr.ReadLine();
            txtTK11.Text = sr.ReadLine();
            txtTK12.Text = sr.ReadLine();
            txtTK13.Text = sr.ReadLine();
            txtTK14.Text = sr.ReadLine();
            txtTK15.Text = sr.ReadLine();
            txtTK16.Text = sr.ReadLine();
            txtTK17.Text = sr.ReadLine();
            txtTK18.Text = sr.ReadLine();
            txtTK19.Text = sr.ReadLine();
            txtTK20.Text = sr.ReadLine();
            txtTK21.Text = sr.ReadLine();
            txtTK22.Text = sr.ReadLine();
            txtTK23.Text = sr.ReadLine();
            txtTK24.Text = sr.ReadLine();


            // učitavaju se konzumi industrijske pare
            txtKP1.Text = sr.ReadLine();
            txtKP2.Text = sr.ReadLine();
            txtKP3.Text = sr.ReadLine();
            txtKP4.Text = sr.ReadLine();
            txtKP5.Text = sr.ReadLine();
            txtKP6.Text = sr.ReadLine();
            txtKP7.Text = sr.ReadLine();
            txtKP8.Text = sr.ReadLine();
            txtKP9.Text = sr.ReadLine();
            txtKP10.Text = sr.ReadLine();
            txtKP11.Text = sr.ReadLine();
            txtKP12.Text = sr.ReadLine();
            txtKP13.Text = sr.ReadLine();
            txtKP14.Text = sr.ReadLine();
            txtKP15.Text = sr.ReadLine();
            txtKP16.Text = sr.ReadLine();
            txtKP17.Text = sr.ReadLine();
            txtKP18.Text = sr.ReadLine();
            txtKP19.Text = sr.ReadLine();
            txtKP20.Text = sr.ReadLine();
            txtKP21.Text = sr.ReadLine();
            txtKP22.Text = sr.ReadLine();
            txtKP23.Text = sr.ReadLine();
            txtKP24.Text = sr.ReadLine();

            // učitavaju se ostali podaci

            txtCijenaElEnergijeNT.Text = sr.ReadLine();
            txtCijenaElEnergijeVT.Text = sr.ReadLine();
            txtCijenaPlin.Text = sr.ReadLine();
            txtCijenaMazut.Text = sr.ReadLine();

            txtCijenaToplina.Text = sr.ReadLine();
            txtCijenaPara.Text = sr.ReadLine();

            txtBrojTurbina.Text = sr.ReadLine();
            txtNapunjenostSpremnika.Text = sr.ReadLine();


            if (sr.ReadLine() == "true")
            {
                cbxKBlok.Checked = true;
            }
            else
            {
                cbxKBlok.Checked = false;
            }


            if (sr.ReadLine() == "true")
            {
                cbxLBlok.Checked = true;
            }
            else
            {
                cbxLBlok.Checked = false;
            }


            if (sr.ReadLine() == "true")
            {
                cbxCBlok.Checked = true;
            }
            else
            {
                cbxCBlok.Checked = false;
            }

            sr.Close();
        }


        //----------------------------------------------------------------------------------------------------------------------  

        private void opcijeToolStripMenuItem_Click(object sender, EventArgs e)
        {
            formaOpcije.Show(this);
        }

        //----------------------------------------------------------------------------------------------------------------------  

        private void btnIconSave_Click(object sender, EventArgs e)
        {
            SaveData();
        }

        //----------------------------------------------------------------------------------------------------------------------  

        private void btnIconOpen_Click(object sender, EventArgs e)
        {
            ReadData();
        }

        //----------------------------------------------------------------------------------------------------------------------  

        private void btnIconImportData_Click(object sender, EventArgs e)
        {
            UnosKonzuma();
        }

        //----------------------------------------------------------------------------------------------------------------------  

        private void btnWriteToExcel_Click(object sender, EventArgs e)
        {
            IspisExcel.IspisivanjeRezultataUExcel(mat);
        }

        //----------------------------------------------------------------------------------------------------------------------  

        private void btnIconShema_Click(object sender, EventArgs e)
        {
            formShema.rezultatiOptimizacije(mat);
            formShema.Show(this);
        }

        //----------------------------------------------------------------------------------------------------------------------  


        private void btnIconOpcije_Click(object sender, EventArgs e)
        {
            formaOpcije.Show(this);
        }

        //----------------------------------------------------------------------------------------------------------------------  

        private void izlazToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Close();
        }




    }

}

