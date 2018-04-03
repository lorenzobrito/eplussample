using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Linq;
using System.IO;
using OfficeOpenXml;

namespace TextEpplus
{
    public partial class Form1 : Form
    {
      public  class DataSource
        {
            public string Column1 { get; set; }
            public string Column2 { get; set; }
            public string Column3 { get; set; }
            public string Column4 { get; set; }
            public string Column5 { get; set; }
        }
        public Form1()
        {
            InitializeComponent();
            dataSource_ = new DataSource();
          
        }
        public DataSource dataSource_;
        public List<DataSource> dataSource
        {
            get;
            set;
           
    }
            

        private void button1_Click(object sender, EventArgs e)
        {
            var lista = this.dataSourceBindingSource.List ;
            int fila = 5;
            char column = 'B';
        
            string FilePath = @"C:\Users\lorenzo.brito\Documents\test.xlsx";
           
            Console.WriteLine();

            FileInfo existingFile = new FileInfo(FilePath);
            Utils.OutputDir = new DirectoryInfo(@"C:\Users\lorenzo.brito\Documents");
            FileInfo newFile = Utils.GetFileInfo("newfilesample.xlsx");
           // FileInfo templateFile = Utils.GetFileInfo(FilePath, false);
            using (ExcelPackage package = new ExcelPackage(newFile, existingFile))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets["Sheet1"];
                foreach (DataSource item in lista)
                {
                    var x = item;
                    
                        worksheet.Cells[column++.ToString() + fila.ToString()].Value = item.Column1;
                        worksheet.Cells[column++.ToString() + fila.ToString()].Value = item.Column2;
                        worksheet.Cells[column++.ToString() + fila.ToString()].Value = item.Column3;
                        worksheet.Cells[column++.ToString() + fila.ToString()].Value = item.Column4;
                        worksheet.Cells[column++.ToString() + fila.ToString()].Value = item.Column5;
                   
                    column = 'B';
                    fila++;
                }
                package.Save();
            }
        }
    }
}
