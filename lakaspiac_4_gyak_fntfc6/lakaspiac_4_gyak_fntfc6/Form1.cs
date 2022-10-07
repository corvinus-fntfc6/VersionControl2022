using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.Reflection;

namespace lakaspiac_4_gyak_fntfc6
{
    public partial class Form1 : Form
    {
        RealEstateEntities context = new RealEstateEntities();
        List<Flat> Flats;
        Excel.Application xlApp;
        Excel.Workbook xlWB;
        Excel.Worksheet xlSheet;

        public Form1()
        {
            InitializeComponent();
            LoadData();
            CreateExcel();

        }

        private void LoadData()
        {
            List<Flat> Flats = context.Flats.ToList();
        }

        private void CreateExcel()
        {
            try
            {
                xlApp = new Excel.Application(); //app betoltese
                xlWB = xlApp.Workbooks.Add(); //uj munkaf
                xlSheet = xlWB.ActiveSheet; //uj munkalap a munkafuzetben
                CreateTable(); //tabla letrehozas

                xlApp.Visible = true;
                xlApp.UserControl = true; //felhasznalo szamara is lathato, majd ő kontrollálja

            }
            catch (Exception ex)
            {
                string errMsg = string.Format("Error: {0\nLine: {1}", ex.Message, ex.Source);
                MessageBox.Show(errMsg, "Error");

                xlWB.Close(false, Type.Missing, Type.Missing);
                xlApp.Quit();
                xlWB = null;
                xlApp = null;
            }
        }

        private void CreateTable()
        {
            string[] headers = new string[]
            {
                "Kód",
                "Eladó",
                "Oldal",
                "Kerület",
                "Lift",
                "Szobák száma",
                "Alapterület (m2)",
                "Ár (mFt)",
                "Négyzetméter ár (Ft/m2)"
            };
            for (int i = 0; i < headers.Length; i++)
            {
                xlSheet.Cells[1, 1] = headers[0];
            }

            object[,] values = new object[Flats.Count, headers.Length];

            int counter = 0;
            foreach (Flat f in Flats)
            {
                values[counter, 0] = f.Code;
                values[counter, 1] = f.Code;
                values[counter, 2] = f.Code;
                values[counter, 3] = f.Code;
                values[counter, 4] = f.Code;
                values[counter, 5] = f.Code;
                values[counter, 6] = f.Code;
                values[counter, 7] = f.Code;
                values[counter, 8] = "";
                counter++;
            }
        }
    }
}
