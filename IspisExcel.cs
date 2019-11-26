using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;

namespace TETOAplikacija
{
    static class IspisExcel
    {
        public static void IspisivanjeRezultataUExcel(double[,] mat)
        {
            Microsoft.Office.Interop.Excel.Application xla = new Microsoft.Office.Interop.Excel.Application();
            Workbook wb = xla.Workbooks.Add(XlSheetType.xlWorksheet);
            Worksheet ws = (Worksheet)xla.ActiveSheet;
            xla.Visible = true;


            int row1 = 0;
            int column1 = 1;
            int odmakStupac1 = 1;

            //L blok
            ws.Cells[row1 + 1, column1 + 0] = "Sat";
            ws.Cells[row1 + 1, column1 + 0 + odmakStupac1] = "opt_L";
            ws.Cells[row1 + 1, column1 + 1 + odmakStupac1] = "P_pt_L, [MW]";
            ws.Cells[row1 + 1, column1 + 2 + odmakStupac1] = "P_ee_L_uk, [MW]";
            ws.Cells[row1 + 1, column1 + 3 + odmakStupac1] = "Q_C5, [MW]";
            ws.Cells[row1 + 1, column1 + 4 + odmakStupac1] = "Q_ZVL, [MW]";
            ws.Cells[row1 + 1, column1 + 5 + odmakStupac1] = "D_10_L, [t/h]";
            ws.Cells[row1 + 1, column1 + 6 + odmakStupac1] = "Dw8, [t/h]";
            ws.Cells[row1 + 1, column1 + 7 + odmakStupac1] = "D_TAL_1ex, [t/h]";
            ws.Cells[row1 + 1, column1 + 8 + odmakStupac1] = "D_TAL_2ex, [t/h]";
            ws.Cells[row1 + 1, column1 + 9 + odmakStupac1] = "Dw9, [t/h]";
            ws.Cells[row1 + 1, column1 + 10 + odmakStupac1] = "gorivoL, [GJ/h]";
            ws.Cells[row1 + 1, column1 + 11 + odmakStupac1] = "Uk. elektr. L, [MJ]";
            ws.Cells[row1 + 1, column1 + 12 + odmakStupac1] = "Q_ukupno L, [MJ]";
            ws.Cells[row1 + 1, column1 + 13 + odmakStupac1] = "Q_korisno L, [MJ]";
            ws.Cells[row1 + 1, column1 + 14 + odmakStupac1] = "gorivoL, [MJ]";
            ws.Cells[row1 + 1, column1 + 15 + odmakStupac1] = "etaL";
            ws.Cells[row1 + 1, column1 + 16 + odmakStupac1] = "eta_eL";
            ws.Cells[row1 + 1, column1 + 17 + odmakStupac1] = "eta_tL";
            ws.Cells[row1 + 1, column1 + 18 + odmakStupac1] = "UPE_L";
            ws.Cells[row1 + 1, column1 + 19 + odmakStupac1] = "ide za C6 i paru, [kg/s]";

            Microsoft.Office.Interop.Excel.Range chartRange1 = ws.get_Range("a1", "u1");
            chartRange1.Interior.Color = XlRgbColor.rgbAqua;
            chartRange1.Font.Bold = true;
            chartRange1.ColumnWidth = 17;


            //K blok

            // odmak od ispisa podataka za K blok
            int odmak = 24 + 4;
            int odmakStupac2 = 2;

            ws.Cells[row1 + 1 + odmak, column1 + 0] = "Sat";
            ws.Cells[row1 + 1 + odmak, column1 + 0 + 1] = "opt_K1";
            ws.Cells[row1 + 1 + odmak, column1 + 0 + 2] = "opt_K2";
            ws.Cells[row1 + 1 + odmak, column1 + 1 + odmakStupac2] = "P_ee_K, [MW]";
            ws.Cells[row1 + 1 + odmak, column1 + 2 + odmakStupac2] = "P_pt_K1, [MW";
            ws.Cells[row1 + 1 + odmak, column1 + 3 + odmakStupac2] = "P_pt_K2, [MW]";
            ws.Cells[row1 + 1 + odmak, column1 + 4 + odmakStupac2] = "Q_C4, [MW]";
            ws.Cells[row1 + 1 + odmak, column1 + 5 + odmakStupac2] = "Q_ZVK1, [MW]";
            ws.Cells[row1 + 1 + odmak, column1 + 6 + odmakStupac2] = "Q_ZVK2, [MW]";
            ws.Cells[row1 + 1 + odmak, column1 + 7 + odmakStupac2] = "D_10_K1, [t/h]";
            ws.Cells[row1 + 1 + odmak, column1 + 8 + odmakStupac2] = "D_w5, [t/h]";
            ws.Cells[row1 + 1 + odmak, column1 + 9 + odmakStupac2] = "D_10_K2, [t/h]";
            ws.Cells[row1 + 1 + odmak, column1 + 10 + odmakStupac2] = "D_w7, [t/h]";
            ws.Cells[row1 + 1 + odmak, column1 + 11 + odmakStupac2] = "D_C4, [t/h]";
            ws.Cells[row1 + 1 + odmak, column1 + 12 + odmakStupac2] = "D_C6, [t/h]";
            ws.Cells[row1 + 1 + odmak, column1 + 13 + odmakStupac2] = "D_91_K1, [t/h]";
            ws.Cells[row1 + 1 + odmak, column1 + 14 + odmakStupac2] = "D_91_K2, [t/h]";
            ws.Cells[row1 + 1 + odmak, column1 + 15 + odmakStupac2] = "Uk. elektr. K, [MJ]";
            ws.Cells[row1 + 1 + odmak, column1 + 16 + odmakStupac2] = "Q_ukupnoK, [MJ]";
            ws.Cells[row1 + 1 + odmak, column1 + 17 + odmakStupac2] = "Q_korisnoK, [MJ]";
            ws.Cells[row1 + 1 + odmak, column1 + 18 + odmakStupac2] = "gorivoK, [MJ]";
            ws.Cells[row1 + 1 + odmak, column1 + 19 + odmakStupac2] = "etaK";
            ws.Cells[row1 + 1 + odmak, column1 + 20 + odmakStupac2] = "eta_eK";
            ws.Cells[row1 + 1 + odmak, column1 + 21 + odmakStupac2] = "eta_tK";
            ws.Cells[row1 + 1 + odmak, column1 + 22 + odmakStupac2] = "UPE_K";

            Microsoft.Office.Interop.Excel.Range chartRange2 = ws.get_Range("a29", "y29");
            chartRange2.Font.Bold = true;
            chartRange2.Interior.Color = XlRgbColor.rgbAqua;




            //C blok

            // odmak od ispisa podataka za C blok
            int odmak1 = odmak + 28;

            ws.Cells[row1 + 1 + odmak1, column1 + 0] = "Sat";
            ws.Cells[row1 + 1 + odmak1, column1 + 1] = "P_ee_C, [MW]";
            ws.Cells[row1 + 1 + odmak1, column1 + 2] = "optC";
            ws.Cells[row1 + 1 + odmak1, column1 + 3] = "gorivoC, [MJ]";
            ws.Cells[row1 + 1 + odmak1, column1 + 4] = "D_140_C, [t/h] ";
            ws.Cells[row1 + 1 + odmak1, column1 + 5] = "D_TAC_In, [t/h]";
            ws.Cells[row1 + 1 + odmak1, column1 + 6] = "Q_C, [MW]";


            Microsoft.Office.Interop.Excel.Range chartRange3 = ws.get_Range("a57", "g57");
            chartRange3.Font.Bold = true;
            chartRange3.Interior.Color = XlRgbColor.rgbAqua;


            // ispis ostalih rezultata
            ws.Cells[86, 1] = "Sat";
            ws.Cells[86, 2] = "Q_vk, [MW]";
            ws.Cells[86, 3] = "Q_C6, [MW]";
            ws.Cells[86, 4] = "Profit, [EUR]";

            Microsoft.Office.Interop.Excel.Range chartRange4 = ws.get_Range("a86", "d86");
            chartRange4.Font.Bold = true;
            chartRange4.Interior.Color = XlRgbColor.rgbAqua;

            // ispis UPE i eta_postrojenja
            ws.Cells[2, 23] = "UPE_L, 24 h";
            ws.Cells[3, 23] = "ETA_L, 24 h";

            ws.Cells[30, 27] = "UPE_K, 24 h";
            ws.Cells[31, 27] = "ETA_K, 24 h";

            Microsoft.Office.Interop.Excel.Range chartRange5 = ws.get_Range("w1", "x1");
            chartRange5.Font.Bold = true;
            chartRange5.ColumnWidth = 10;


            Microsoft.Office.Interop.Excel.Range chartRange6 = ws.get_Range("aa1", "ab1");
            chartRange6.Font.Bold = true;
            chartRange6.ColumnWidth = 10;


            //ispis podataka za L blok

            for (int i = 0; i < 24; i++)
            {
                ws.Cells[row1 + i + 2, column1 + 0] = i + 1;
                ws.Cells[row1 + i + 2, column1 + 0 + odmakStupac1] = Math.Round(mat[i, 5 - 1],3);
                ws.Cells[row1 + i + 2, column1 + 1 + odmakStupac1] = Math.Round(mat[i, 21 - 1], 3);
                ws.Cells[row1 + i + 2, column1 + 2 + odmakStupac1] = Math.Round(mat[i, 31 - 1], 3);
                ws.Cells[row1 + i + 2, column1 + 3 + odmakStupac1] = Math.Round(mat[i, 44 - 1], 3);
                ws.Cells[row1 + i + 2, column1 + 4 + odmakStupac1] = Math.Round(mat[i, 22 - 1], 3);
                ws.Cells[row1 + i + 2, column1 + 5 + odmakStupac1] = Math.Round(mat[i, 19 - 1], 3);
                ws.Cells[row1 + i + 2, column1 + 6 + odmakStupac1] = Math.Round(mat[i, 24 - 1], 3);
                ws.Cells[row1 + i + 2, column1 + 7 + odmakStupac1] = Math.Round(mat[i, 6 - 1], 3);
                ws.Cells[row1 + i + 2, column1 + 8 + odmakStupac1] = Math.Round(mat[i, 7 - 1], 3);
                ws.Cells[row1 + i + 2, column1 + 9 + odmakStupac1] = Math.Round(mat[i, 23 - 1], 3);
                ws.Cells[row1 + i + 2, column1 + 10 + odmakStupac1] = Math.Round(mat[i, 18 - 1], 3);
                ws.Cells[row1 + i + 2, column1 + 11 + odmakStupac1] = Math.Round((mat[i, 21 - 1] + mat[i, 31 - 1]) * 3600, 3);
                ws.Cells[row1 + i + 2, column1 + 12 + odmakStupac1] = Math.Round(mat[i, 76 - 1], 3);
                ws.Cells[row1 + i + 2, column1 + 13 + odmakStupac1] = Math.Round(mat[i, 77 - 1], 3);
                ws.Cells[row1 + i + 2, column1 + 14 + odmakStupac1] = Math.Round(mat[i, 18 - 1] * 1000, 3);
                ws.Cells[row1 + i + 2, column1 + 15 + odmakStupac1] = Math.Round(((mat[i, 21 - 1] + mat[i, 31 - 1]) * 3600 + mat[i, 76 - 1]) / (mat[i, 18 - 1] * 1000), 3);
                ws.Cells[row1 + i + 2, column1 + 16 + odmakStupac1] = Math.Round((mat[i, 21 - 1] + mat[i, 31 - 1]) / mat[i, 18 - 1], 3);
                ws.Cells[row1 + i + 2, column1 + 17 + odmakStupac1] = Math.Round(mat[i, 77 - 1] / (mat[i, 18 - 1] * 3600), 3);
                ws.Cells[row1 + i + 2, column1 + 18 + odmakStupac1] = Math.Round(mat[i, 79 - 1] / 10000, 3);
                ws.Cells[row1 + i + 2, column1 + 19 + odmakStupac1] = Math.Round(mat[i, 51 - 1], 3);
            }


            // ispis podataka za K blok
            for (int i = 0; i < 24; i++)
            {
                ws.Cells[row1 + i + 2 + odmak, column1 + 0] = i + 1;
                ws.Cells[row1 + i + 2 + odmak, column1 + 1] = Math.Round(mat[i, 2 - 1],3);
                ws.Cells[row1 + i + 2 + odmak, column1 + 2] = Math.Round(mat[i, 90 - 1],3);
                ws.Cells[row1 + i + 2 + odmak, column1 + 1 + odmakStupac2] = Math.Round(mat[i, 32 - 1],3);
                ws.Cells[row1 + i + 2 + odmak, column1 + 2 + odmakStupac2] = Math.Round(mat[i, 12 - 1], 3);
                ws.Cells[row1 + i + 2 + odmak, column1 + 3 + odmakStupac2] = Math.Round(mat[i, 16 - 1], 3);
                ws.Cells[row1 + i + 2 + odmak, column1 + 4 + odmakStupac2] = Math.Round(mat[i, 29 - 1], 3);
                ws.Cells[row1 + i + 2 + odmak, column1 + 5 + odmakStupac2] = Math.Round(mat[i, 13 - 1], 3);
                ws.Cells[row1 + i + 2 + odmak, column1 + 6 + odmakStupac2] = Math.Round(mat[i, 17 - 1], 3);
                ws.Cells[row1 + i + 2 + odmak, column1 + 7 + odmakStupac2] = Math.Round(mat[i, 10 - 1], 3);
                ws.Cells[row1 + i + 2 + odmak, column1 + 8 + odmakStupac2] = Math.Round(mat[i, 27 - 1], 3);
                ws.Cells[row1 + i + 2 + odmak, column1 + 9 + odmakStupac2] = Math.Round(mat[i, 14 - 1], 3);
                ws.Cells[row1 + i + 2 + odmak, column1 + 10 + odmakStupac2] = Math.Round(mat[i, 25 - 1], 3);
                ws.Cells[row1 + i + 2 + odmak, column1 + 11 + odmakStupac2] = Math.Round(mat[i, 41 - 1], 3);
                ws.Cells[row1 + i + 2 + odmak, column1 + 12 + odmakStupac2] = Math.Round(mat[i, 43 - 1], 3);
                ws.Cells[row1 + i + 2 + odmak, column1 + 13 + odmakStupac2] = Math.Round(mat[i, 11 - 1], 3);
                ws.Cells[row1 + i + 2 + odmak, column1 + 14 + odmakStupac2] = Math.Round(mat[i, 15 - 1], 3);
                ws.Cells[row1 + i + 2 + odmak, column1 + 15 + odmakStupac2] = Math.Round(mat[i, 80 - 1], 3);
                ws.Cells[row1 + i + 2 + odmak, column1 + 16 + odmakStupac2] = Math.Round(mat[i, 86 - 1], 3);
                ws.Cells[row1 + i + 2 + odmak, column1 + 17 + odmakStupac2] = Math.Round(mat[i, 87 - 1], 3);
                ws.Cells[row1 + i + 2 + odmak, column1 + 18 + odmakStupac2] = Math.Round(mat[i, 88 - 1], 3);
                ws.Cells[row1 + i + 2 + odmak, column1 + 19 + odmakStupac2] = Math.Round((mat[i, 80 - 1] + mat[i, 86 - 1]) / mat[i, 88 - 1], 3);
                ws.Cells[row1 + i + 2 + odmak, column1 + 20 + odmakStupac2] = Math.Round(mat[i, 80 - 1] / mat[i, 88 - 1], 3);
                ws.Cells[row1 + i + 2 + odmak, column1 + 21 + odmakStupac2] = Math.Round(mat[i, 87 - 1] / mat[i, 88 - 1], 3);
                ws.Cells[row1 + i + 2 + odmak, column1 + 22 + odmakStupac2] = Math.Round(mat[i, 89 - 1] / 10000, 3);
            }

            // ispis podataka za C blok

            for (int i = 0; i < 24; i++)
            {
                ws.Cells[row1 + i + 2 + odmak1, column1 + 0] = i + 1;
                ws.Cells[row1 + i + 2 + odmak1, column1 + 1] = Math.Round(mat[i, 66 - 1], 3);
                ws.Cells[row1 + i + 2 + odmak1, column1 + 2] = Math.Round(mat[i, 62 - 1], 3);
                ws.Cells[row1 + i + 2 + odmak1, column1 + 3] = Math.Round(mat[i, 67 - 1], 3);
                ws.Cells[row1 + i + 2 + odmak1, column1 + 4] = Math.Round(mat[i, 65 - 1], 3);
                ws.Cells[row1 + i + 2 + odmak1, column1 + 5] = Math.Round(mat[i, 64 - 1], 3);
                ws.Cells[row1 + i + 2 + odmak1, column1 + 6] = Math.Round(mat[i, 63 - 1], 3);
            }


            // ispis ostalih rezultata

            for (int i = 0; i < 24; i++)
            {
                ws.Cells[87 + i, 1] = i + 1;
                ws.Cells[87 + i, 2] = Math.Round(mat[i, 36 - 1], 3);
                ws.Cells[87 + i, 3] = Math.Round(mat[i, 45 - 1], 3);
                ws.Cells[87 + i, 4] = Math.Round(mat[i, 113 - 1], 3);

            }

            // ispis UPE i eta_postrojenja
            ws.Cells[2, 24] = Math.Round(mat[1, 101 - 1] / 100, 3);
            ws.Cells[3, 24] = Math.Round(mat[1, 102 - 1] / 100, 3);

            ws.Cells[30, 28] = Math.Round(mat[1, 106 - 1] / 10000, 3);
            ws.Cells[31, 28] = Math.Round(mat[1, 107 - 1], 3);
        }
    }
}
