using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;


namespace M7_P05
{
    public partial class Form1 : Form
    {
        Articulo[] articulos;
        TextBox[] cajasPreciosUni;
        TextBox[] cajasPrecios;
        ComboBox[] combosArticulos;
        NumericUpDown[] spinners;
        ComboBox combo;
        
        Word.Application apWord = null;
        Word.Document docWord = null;

        Excel.Application appExcel = null;
        Excel.Workbook workBookExcel = null;
        Excel.Worksheet workSheetExcel = null;

        public Form1()
        {
            InitializeComponent();
            inicio();
        }
        private void inicio() {
            Articulo art1 = new Articulo("HDJ-500", 79.95f);
            Articulo art2 = new Articulo("HDJ-700", 127);
            Articulo art3 = new Articulo("HDJ-1500", 157);
            Articulo art4 = new Articulo("HDJ-c70", 169);
            Articulo art5 = new Articulo("HDJ-2000MK2", 347);
            Articulo art6 = new Articulo("S-DJ50X", 159);
            Articulo art7 = new Articulo("S-DJ80X", 179);
            Articulo art8 = new Articulo("DM-40", 147);
            Articulo art9 = new Articulo("XDJ-1000MK2", 1299);
            Articulo art10 = new Articulo("CDJ-2000NXS2", 2255);
            articulos = new Articulo[10];
            articulos[0] = art1;
            articulos[1] = art2;
            articulos[2] = art3;
            articulos[3] = art4;
            articulos[4] = art5;
            articulos[5] = art6;
            articulos[6] = art7;
            articulos[7] = art8;
            articulos[8] = art9;
            articulos[9] = art10;
            
            cajasPreciosUni = new TextBox[] 
            { textBox18, textBox19, textBox21, textBox20, textBox22, textBox23, textBox24};
            
            cajasPrecios = new TextBox[]
            { textBox11, textBox12, textBox13, textBox14, textBox15, textBox16, textBox17};
            
            spinners = new NumericUpDown[]
            { numericUpDown1, numericUpDown2, numericUpDown3, numericUpDown4, numericUpDown5, numericUpDown6, numericUpDown7};

            combosArticulos = new ComboBox[] {c1, c2, c3, c4, c5, c6, c7};

            textBox10.Text = DateTime.Now.ToString("dd/MM/yyyy");





        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            combo = (ComboBox)sender;
            String select = combo.GetItemText(combo.SelectedItem);
           
            for (int i = 0; i< articulos.Length; i++) {
                if(select.Equals(articulos[i].gsNombre)) {
                    float precioUni = articulos[i].gsPrecio;
                    if (combo.Name.Equals("c1")){ 
                        cajasPreciosUni[0].Text = Convert.ToString(precioUni);
                        cajasPrecios[0].Text = Convert.ToString(precioUni * Convert.ToInt32(spinners[0].Value));
                    }
                    else if (combo.Name.Equals("c2")){ 
                        cajasPreciosUni[1].Text = Convert.ToString(precioUni);
                        cajasPrecios[1].Text = Convert.ToString(precioUni * Convert.ToInt32(spinners[1].Value));
                    }
                    else if (combo.Name.Equals("c3")){ 
                        cajasPreciosUni[2].Text = Convert.ToString(articulos[i].gsPrecio);
                        cajasPrecios[2].Text = Convert.ToString(precioUni * Convert.ToInt32(spinners[2].Value));
                    }
                    else if (combo.Name.Equals("c4")){ 
                        cajasPreciosUni[3].Text = Convert.ToString(articulos[i].gsPrecio);
                        cajasPrecios[3].Text = Convert.ToString(precioUni * Convert.ToInt32(spinners[3].Value));
                    }
                    else if (combo.Name.Equals("c5")){ 
                        cajasPreciosUni[4].Text = Convert.ToString(articulos[i].gsPrecio);
                        cajasPrecios[4].Text = Convert.ToString(precioUni * Convert.ToInt32(spinners[4].Value));
                    }
                    else if (combo.Name.Equals("c6")){ 
                        cajasPreciosUni[5].Text = Convert.ToString(articulos[i].gsPrecio);
                        cajasPrecios[5].Text = Convert.ToString(precioUni * Convert.ToInt32(spinners[5].Value));
                    }
                    else if (combo.Name.Equals("c7")){ 
                        cajasPreciosUni[6].Text = Convert.ToString(articulos[i].gsPrecio);
                        cajasPrecios[6].Text = Convert.ToString(precioUni * Convert.ToInt32(spinners[6].Value));
                    }
                }
            }
            calculoTotal();
        }

        private void numericUpDown1_ValueChanged(object sender, EventArgs e)
        {
            NumericUpDown spinner = (NumericUpDown)sender;
            try {
                if (spinner == spinners[0])
                {
                    cajasPrecios[0].Text = Convert.ToString(Convert.ToDouble(cajasPreciosUni[0].Text) * Convert.ToInt32(spinner.Value));
                }
                else if (spinner == spinners[1])
                {
                    cajasPrecios[1].Text = Convert.ToString(Convert.ToDouble(cajasPreciosUni[1].Text) * Convert.ToInt32(spinner.Value));
                }
                else if (spinner == spinners[2])
                {
                    cajasPrecios[2].Text = Convert.ToString(Convert.ToDouble(cajasPreciosUni[2].Text) * Convert.ToInt32(spinner.Value));
                }
                else if (spinner == spinners[3])
                {
                    cajasPrecios[3].Text = Convert.ToString(Convert.ToDouble(cajasPreciosUni[3].Text) * Convert.ToInt32(spinner.Value));
                }
                else if (spinner == spinners[4])
                {
                    cajasPrecios[4].Text = Convert.ToString(Convert.ToDouble(cajasPreciosUni[4].Text) * Convert.ToInt32(spinner.Value));
                }
                else if (spinner == spinners[5])
                {
                    cajasPrecios[5].Text = Convert.ToString(Convert.ToDouble(cajasPreciosUni[5].Text) * Convert.ToInt32(spinner.Value));
                }
                else if (spinner == spinners[6])
                {
                    cajasPrecios[6].Text = Convert.ToString(Convert.ToDouble(cajasPreciosUni[6].Text) * Convert.ToInt32(spinner.Value));
                }

                calculoTotal();
            }
            catch {
                Console.WriteLine("error");
            }
            calculoTotal();


        }
        private void calculoTotal() {
            double total = 0;
            double totalIva = 0;
            for (int i = 0; i< cajasPrecios.Length; i++) {
                try {
                    total = total +  Convert.ToDouble(cajasPrecios[i].Text);
                    textBox26.Text = Convert.ToString(total);
         
                }
                catch {
                    Console.WriteLine("error");
                } 
            }
            try {
                double iva = Convert.ToDouble(comboBox31.SelectedItem.ToString());
                totalIva = total * (1 + (iva / 100));
                textBox25.Text = Convert.ToString(totalIva);
            }
            catch {
                Console.WriteLine("error");
            }
           
        }

        private void CrearFicheroWord()
        {
            Word.Application apWord0 = new Word.Application();
            Word.Document docWord0 = new Word.Document();
            string path = Directory.GetCurrentDirectory();
            path = path + "\\Facturas\\";
            string wordFileNameIn = "plantillaWord.docx";
            string wordFileNameOut = "Factura_out.docx";
            docWord0 = apWord0.Documents.Open(path + wordFileNameIn);
            docWord0.SaveAs(path + wordFileNameOut);
            docWord0.Close();
            apWord0.Quit();
            apWord = new Word.Application();
            docWord = new Word.Document();
            docWord = apWord.Documents.Open(path + wordFileNameOut);
        }

        private void EscribirEnFicheroWord()
        {
            Object bookMarcName; 
            string text = " ";
            
            bookMarcName = "razonSocial";
            text = textBox6.Text;
            docWord.Bookmarks[ref bookMarcName].Select();
            apWord.Selection.TypeText(Text: text);

            bookMarcName = "NIF";
            text = textBox5.Text;
            docWord.Bookmarks[ref bookMarcName].Select();
            apWord.Selection.TypeText(Text: text);

            bookMarcName = "domicilio";
            text = textBox4.Text;
            docWord.Bookmarks[ref bookMarcName].Select();
            apWord.Selection.TypeText(Text: text);

            bookMarcName = "albaran";
            text = textBox8.Text;
            docWord.Bookmarks[ref bookMarcName].Select();
            apWord.Selection.TypeText(Text: text);
            
            bookMarcName = "numPedido";
            text = textBox7.Text;
            docWord.Bookmarks[ref bookMarcName].Select();
            apWord.Selection.TypeText(Text: text);

            bookMarcName = "num";
            text = textBox9.Text;
            docWord.Bookmarks[ref bookMarcName].Select();
            apWord.Selection.TypeText(Text: text);

            bookMarcName = "fecha";
            text = textBox10.Text;
            docWord.Bookmarks[ref bookMarcName].Select();
            apWord.Selection.TypeText(Text: text);

            foreach(TextBox tb in cajasPreciosUni) {
                bookMarcName = tb.Name;
                text = tb.Text;
                docWord.Bookmarks[ref bookMarcName].Select();
                apWord.Selection.TypeText(Text: text);
            }

            foreach (TextBox tb in cajasPrecios)
            {
                bookMarcName = tb.Name;
                text = tb.Text;
                docWord.Bookmarks[ref bookMarcName].Select();
                apWord.Selection.TypeText(Text: text);
            }

            foreach (NumericUpDown num in spinners)
            {
                bookMarcName = num.Name;
                text = Convert.ToString(num.Value);
                docWord.Bookmarks[ref bookMarcName].Select();
                apWord.Selection.TypeText(Text: text);
            }

            foreach (ComboBox cb in combosArticulos)
            {
                bookMarcName = cb.Name;
                text = cb.Text;
                docWord.Bookmarks[ref bookMarcName].Select();
                apWord.Selection.TypeText(Text: text);
            }

            bookMarcName = "textBox26";
            text = textBox26.Text;
            docWord.Bookmarks[ref bookMarcName].Select();
            apWord.Selection.TypeText(Text: text);

            bookMarcName = "c31";
            text = comboBox31.Text;
            docWord.Bookmarks[ref bookMarcName].Select();
            apWord.Selection.TypeText(Text: text);

            bookMarcName = "textBox25";
            text = textBox25.Text;
            docWord.Bookmarks[ref bookMarcName].Select();
            apWord.Selection.TypeText(Text: text);

        }

        private void GuardarFicheroWord()
        {
            docWord.Save();
        }

        private void CerrarWord()
        {
            apWord.Quit();
            System.Runtime.InteropServices.Marshal.FinalReleaseComObject(apWord);
            apWord = null;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            CrearFicheroWord();
            EscribirEnFicheroWord();
            GuardarFicheroWord();
            CerrarWord();
            MessageBox.Show("Word file created");
            string path = @"C:\Users\jordimasmer\Documents\Visual Studio 2015\Projects\M7-P05\bin\Debug\Facturas";
            System.Diagnostics.Process.Start(path);
        }

        private void comboBox31_SelectedIndexChanged(object sender, EventArgs e)
        {
            calculoTotal();
        }


        private void gestionExcel()
        {
            abrirFicheroExcel();
            EscribirEnFicheroExcel();
            guardarFicheroExcel();
            cerrarExcel();
            MessageBox.Show("Excel file created");
            string path = @"C:\Users\jordimasmer\Documents\Visual Studio 2015\Projects\M7-P05\bin\Debug\Facturas";
            System.Diagnostics.Process.Start(path);
        }
        private void abrirFicheroExcel()
        {
            appExcel = new Excel.Application();
            String path = Directory.GetCurrentDirectory() + "\\Facturas\\";
            String ExcelFileNameIn = "plantillaExcel.xlsx";
            workBookExcel = appExcel.Workbooks.Open(path + ExcelFileNameIn);
            workSheetExcel = (Excel.Worksheet)workBookExcel.Worksheets.get_Item(1);
        }
        private void EscribirEnFicheroExcel()
        {
            int fila, col; string text;
            
            fila = 1; col = 2; 
            text = textBox9.Text;
            appExcel.DisplayAlerts = false;
            workSheetExcel.Cells[fila, col] = text;

            fila = 2; col = 2;
            text = textBox10.Text;
            appExcel.DisplayAlerts = false;
            workSheetExcel.Cells[fila, col] = text;

            fila = 3; col = 2;
            text = textBox6.Text;
            appExcel.DisplayAlerts = false;
            workSheetExcel.Cells[fila, col] = text;

            fila = 4; col = 2;
            text = textBox5.Text;
            appExcel.DisplayAlerts = false;
            workSheetExcel.Cells[fila, col] = text;
            
            fila = 5; col = 2;
            text = textBox4.Text;
            appExcel.DisplayAlerts = false;
            workSheetExcel.Cells[fila, col] = text;

            fila = 6; col = 2;
            text = textBox8.Text;
            appExcel.DisplayAlerts = false;
            workSheetExcel.Cells[fila, col] = text;

            fila = 7; col = 2;
            text = textBox7.Text;
            appExcel.DisplayAlerts = false;
            workSheetExcel.Cells[fila, col] = text;

            fila = 10;

            foreach (TextBox tb in cajasPreciosUni)
            {
                col = 3;
                text = tb.Text;
                appExcel.DisplayAlerts = false;
                workSheetExcel.Cells[fila, col] = text;
                fila++;
            }
            fila = 10;
            foreach (TextBox tb in cajasPrecios)
            {
                 col = 4;
                text = tb.Text;
                appExcel.DisplayAlerts = false;
                workSheetExcel.Cells[fila, col] = text;
                fila++;
            }
            fila = 10;
            foreach (NumericUpDown num in spinners)
            {
                col = 2;
                text = num.Text;
                appExcel.DisplayAlerts = false;
                workSheetExcel.Cells[fila, col] = text;
                fila++;
            }
            fila = 10;
            foreach (ComboBox cb in combosArticulos)
            {
                 col = 1;
                text = cb.Text;
                appExcel.DisplayAlerts = false;
                workSheetExcel.Cells[fila, col] = text;
                fila++;
            }

            fila = 17; col = 2;
            text = textBox26.Text;
            appExcel.DisplayAlerts = false;
            workSheetExcel.Cells[fila, col] = text;

            fila = 18; col = 2;
            text = comboBox31.Text;
            appExcel.DisplayAlerts = false;
            workSheetExcel.Cells[fila, col] = text;

            fila = 19; col = 2;
            text = textBox25.Text;
            appExcel.DisplayAlerts = false;
            workSheetExcel.Cells[fila, col] = text;

        }

        private void guardarFicheroExcel()
        {
            String path = Directory.GetCurrentDirectory() + "\\Facturas\\";
            String ExcelFileNameOut = "Factura_out.xlsx";
            workBookExcel.SaveAs(path + ExcelFileNameOut);
        }
        private void cerrarExcel()
        {
            workBookExcel.Close(true);//guardar los cambios:true
            appExcel.Quit();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(workSheetExcel);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(workBookExcel);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(appExcel);
        }

        private void guardarFicheroPDF()
        {
            Word.Application apWord = new Word.Application();
            Word.Document docWord = new Word.Document();
            string path = Directory.GetCurrentDirectory();
            path = path + "\\Facturas\\";
            string wordFileNameIn = "Factura_out.docx";
            docWord = apWord.Documents.Open(path + wordFileNameIn);
            string PdfFileNameOut = path + "Factura_out.pdf";
            docWord.SaveAs(PdfFileNameOut, Word.WdSaveFormat.wdFormatPDF);
            docWord.Close();
            apWord.Quit();
        }

        private void button2_Click(object sender, EventArgs e)
        {
        try {
            guardarFicheroPDF();
            MessageBox.Show("PDF file created");
            string path = @"C:\Users\jordimasmer\Documents\Visual Studio 2015\Projects\M7-P05\bin\Debug\Facturas";
            System.Diagnostics.Process.Start(path);
            }
        catch {
                MessageBox.Show("PDF file could not be created");
            }    
        
        }

        private void button3_Click(object sender, EventArgs e)
        {
            try
            {
                gestionExcel();
                MessageBox.Show("Excel file created");
                string path = @"C:\Users\jordimasmer\Documents\Visual Studio 2015\Projects\M7-P05\bin\Debug\Facturas";
                System.Diagnostics.Process.Start(path);
            }
            catch
            {
                MessageBox.Show("Excel file could not be created");
            }

            
        }
    }
}
