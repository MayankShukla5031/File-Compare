using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Microsoft.Win32;
using System.Windows.Media.Animation;
using System.IO;
using System.Reflection;
//using Microsoft.Office;

namespace File_Compare
{
    using System.Collections;
    using System.Threading;

    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    /// 
    using Excel = Microsoft.Office.Interop.Excel;

    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
            WindowStartupLocation = WindowStartupLocation.CenterScreen;
            Desktop_path = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
           
        }

         
        string file1_path, file2_path, Desktop_path, file1_text, file2_text;
        int column, row, TotalColumn,file_type;
        int[] validColumn = new int[150];
        Hashtable hashtable = new Hashtable();
        List<string> DuplicateValues = new List<string>();

        //file1
        Excel.Application excelApp1 = new Excel.Application();
        Excel.Workbook workbook1;
        Excel.Worksheet worksheet1;
        Excel.Range range1;

        //file2
        Excel.Application excelApp2 = new Excel.Application();
        Excel.Workbook workbook2;
        Excel.Worksheet worksheet2;
        Excel.Range range2;
        
        //file3
        Excel.Application newapp = new Excel.Application();
        Excel.Workbook new_workbook;
        Excel.Worksheet new_worksheet;
        Excel.Range new_workSheet_range;

        private void BrowseButton1_MouseUp(object sender, MouseButtonEventArgs e)
        {
            browsefile1.Margin = new Thickness(0);
            var filereader = new OpenFileDialog();
          //  filereader.Filter = "Excel Files (.xls)|*.xlsx | Excel Files (.xls)|*.xls";
            filereader.ShowDialog();
            file1_path = filereader.FileName;
            file1_path = file1_path.Replace("\\", "/");
            file1.Text = file1_path.Substring(file1_path.LastIndexOf("/") + 1);
         }

        private void BrowseButton2_MouseUp(object sender, MouseButtonEventArgs e)
        {
            browsefile2.Margin = new Thickness(0);
            var filereader = new OpenFileDialog();
            // filereader.Filter = "Excel Files";
            filereader.ShowDialog();
            file2_path = filereader.FileName;
            file2_path = file2_path.Replace("\\", "/");
            file2.Text = file2_path.Substring(file2_path.LastIndexOf("/") + 1);             
        }

        private void CompareButton_MouseUp(object sender, MouseButtonEventArgs e)
        {
            CompareButton.Margin = new Thickness(0);
            if (file1.Text == "" || file2.Text == "" || NewFileName.Text == "")
             {
                errorMessage.Content = "Kindly Fill All the Feilds...";
                AnimateErrorGrid();
                return;
             }

            if (row == 0)
            {
                errorMessage.Content = "Kindly Select a File Type...";
                AnimateErrorGrid();
                return;
            }

            // creating folder on desktop
            if (!Directory.Exists(Desktop_path + "\\Bhaiya"))
                Directory.CreateDirectory(Desktop_path + "\\Bhaiya");

            string str = Desktop_path + "\\Bhaiya\\" + NewFileName.Text+".xls";

            if (File.Exists(str))
            {
                errorMessage.Content = "File name already exist";
                AnimateErrorGrid();
                return;
            }

            //Opening first File
            workbook1 = excelApp1.Workbooks.Open(file1_path);
            worksheet1 = (Excel.Worksheet)workbook1.Sheets[1];
            range1 = worksheet1.UsedRange;


            //openinng Second File
            workbook2 = excelApp2.Workbooks.Open(file2_path);
            worksheet2 = (Excel.Worksheet)workbook2.Sheets[1];
            range2 = worksheet2.UsedRange;
           
            
            // Creating New File
            newapp = new Excel.Application();
            newapp.Visible = true;
            new_workbook = newapp.Workbooks.Add(1);
            new_worksheet = (Excel.Worksheet)new_workbook.Sheets[1];
            new_workSheet_range = worksheet2.UsedRange;
            object misValue = System.Reflection.Missing.Value;

            TotalColumn = range2.Columns.Count;

           

            // saving the excel file in Bhaiya folder 
            new_workbook.SaveAs(Desktop_path + "\\Bhaiya\\" + NewFileName.Text, Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);

           
            DoHashingofSecondFile();

            WriteDatainResultFile();

            RearrangeNewSheet();

            AddDuplicates();

            CloseAllSheet();

        }

        private void DoHashingofSecondFile()
        {
            try
            {
                string value;
                int row_count = range2.Rows.Count;
                for (int localrow = row + 1; localrow <= row_count; localrow++)
                {
                    value = Convert.ToString(worksheet2.Cells[localrow, column].Value2) + Convert.ToString(worksheet2.Cells[localrow, column + 8].Value2);
                    if (!hashtable.ContainsKey(value))
                    {
                        hashtable.Add(value, localrow);
                      //  Console.WriteLine(localrow);
                    }
                    else
                        DuplicateValues.Add(value);
                }
            }
            catch(Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }

        private void WriteDatainResultFile()
        {
            try
            {
                int local_row_sheet1, local_row_sheet2, local_column,new_file_column, flag;
                new_file_column = 1;
                string column_name;

                for (local_column = column; local_column <= TotalColumn; local_column++)
                {
                    if (validColumn[local_column] !=0)
                    {
                        file2_text = Convert.ToString(worksheet2.Cells[row, local_column].Value2);
                        new_worksheet.Cells[row, new_file_column] = file2_text;
                        new_worksheet.Cells[row, new_file_column].Interior.Color = System.Drawing.Color.Green.ToArgb();
                        validColumn[local_column] = new_file_column++;
                        // new_worksheet.Cells[row, column].Interior.Color = System.Drawing.Color.Green.ToArgb();
                    }
                }


                int row_count = range1.Rows.Count; 

                for (local_row_sheet1 = row + 1; local_row_sheet1 <= row_count; local_row_sheet1++)
                {
                    flag = 0;
                    new_file_column = 2;
                    column_name = file1_text = Convert.ToString(worksheet1.Cells[local_row_sheet1, column].Value2)+ Convert.ToString(worksheet1.Cells[local_row_sheet1, column + 8].Value2);

                    if (hashtable.ContainsKey(file1_text))
                    {
                        local_row_sheet2 = (int)(hashtable[file1_text]);
                        hashtable.Remove(file1_text);
                       // Console.WriteLine("rowin2 : " + local_row_sheet2);
                    }
                    else
                        continue;

                    for (local_column = column + 1; local_column <= TotalColumn; local_column++)
                    {
                        if (validColumn[local_column] != 0)
                        {
                            file1_text = Convert.ToString(worksheet1.Cells[local_row_sheet1, local_column].Value2);
                            file2_text = Convert.ToString(worksheet2.Cells[local_row_sheet2, local_column].Value2);

                            if (file1_text != file2_text)
                            {
                                if ((file_type==1 && local_column == 67) || (file_type==2 && (local_column == 34 || local_column == 35)))
                                {
                                    file1_text = file1_text.Replace('-','0');
                                    string str = file2_text.Replace('-','0');
                                    double value1;
                                    double.TryParse(file1_text, out value1);
                                    double value2;
                                    double.TryParse(str, out value2);
                                    if ((value2 - value1) > 3 || value2 == 0 || value1 == 0)
                                    {
                                        flag = 1;
                                        new_worksheet.Cells[local_row_sheet2, validColumn[local_column]] = file2_text;
                                        new_worksheet.Cells[local_row_sheet2, validColumn[local_column]].Interior.Color = System.Drawing.Color.FromArgb(255, 106, 90, 205);
                                    }
                                }
                                else
                                {
                                    flag = 1;
                                    new_worksheet.Cells[local_row_sheet2, validColumn[local_column]] = file2_text;
                                    new_worksheet.Cells[local_row_sheet2, validColumn[local_column]].Interior.Color = System.Drawing.Color.FromArgb(255, 106, 90, 205);
                                }
                            }
                        }
                    }
                    if (flag == 1)
                    {
                        new_worksheet.Cells[local_row_sheet2, 1] = column_name;
                    }
                }

          //      new_file_column = 1;
                if (hashtable.Count > 0)
                {
                    foreach (DictionaryEntry value in hashtable)
                    {
                        new_file_column = 1;
                        for (local_column = 1; local_column <= TotalColumn; local_column++)
                        {
                            if (validColumn[local_column] !=0)
                            {
                                file2_text = Convert.ToString(worksheet2.Cells[(int)value.Value, local_column].Value2);
                                new_worksheet.Cells[value.Value, new_file_column] = file2_text;
                                new_worksheet.Cells[value.Value, new_file_column++].Interior.Color = System.Drawing.Color.FromArgb(255, 106, 90, 205);
                            }
                        }
                    }
                }       

            }
            catch(Exception ex)
            {
                Console.WriteLine(ex.Message);
            }


        }

        int new_row_count;
        private void RearrangeNewSheet()
        {
            try
            {
                new_row_count = new_workSheet_range.Rows.Count;
                for (int local_row_sheet = 1; local_row_sheet <= new_row_count; local_row_sheet++)
                {
                    String str = Convert.ToString(new_worksheet.Cells[local_row_sheet, 1].Value2);
                    if (str == "" || str == null)
                    {
                        ((Excel.Range)new_worksheet.Rows[local_row_sheet]).Delete(Excel.XlDirection.xlUp);
                        new_row_count--;
                        local_row_sheet--;
                    } 
                }
            }
            catch(Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }

        private void AddDuplicates()
        {
            try
            {
                int local_rows = new_row_count + 1;
                if (DuplicateValues.Count > 0)
                {
                    foreach (String value in DuplicateValues)
                    {
                        new_worksheet.Cells[local_rows, 1] = value;
                        new_worksheet.Cells[local_rows, 2] = "Duplicate";
                        new_worksheet.Cells[local_rows, 1].Interior.Color = new_worksheet.Cells[local_rows, 2].Interior.Color = System.Drawing.Color.FromArgb(255, 106, 90, 205);
                        local_rows++;
                    }
                }
            }
            catch(Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }

        private void CloseAllSheet()
        {
            try
            {
                workbook1.Close(true, Missing.Value, Missing.Value);
                excelApp1.Quit();
                workbook2.Close(true, Missing.Value, Missing.Value);
                excelApp2.Quit();
                
                releaseObject(excelApp1);
                releaseObject(excelApp2);
                releaseObject(workbook1);
                releaseObject(workbook2);

               
              
                new_workbook.Save();
                new_workbook.Close(true, Missing.Value, Missing.Value);
                newapp.Quit();
                releaseObject(newapp);
                releaseObject(new_workbook);

                Application.Current.Shutdown();
            }
            catch(Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }

        private void releaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
                MessageBox.Show("Unable to release the Object " + ex.ToString());
            }
        }

        private void AnimateErrorGrid()
        {
            DoubleAnimation showanimation = new DoubleAnimation(0d,1d,TimeSpan.FromMilliseconds(500));
            showanimation.Completed += showanimation_Completed;
            error_msg_border.BeginAnimation(OpacityProperty,showanimation);
        }     

        private void ComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            // for (int i = 0; i < 100; i++)
            //     Console.WriteLine(validColumn[i]);

            validColumn = Enumerable.Repeat(0,150).ToArray();

            if (FileType.SelectedIndex == 1)
            {
                file_type = 1;
                row = 1;
                column = 5;
                validColumn[5] = validColumn[6] = validColumn[8] = validColumn[10] = validColumn[34] = validColumn[35] = validColumn[50] = validColumn[61] = validColumn[65] = validColumn[67] = validColumn[98] = validColumn[99] = 1;
                // f6 h8 j10 ah34 ai35 ax50 bi61 bm65 bo67 ct98 cu99
            }
            else if (FileType.SelectedIndex == 2)
            {
                file_type = 2;
                row = 8;
                column = 1;
                validColumn[1] = validColumn[27] = validColumn[34] = validColumn[35] = 1;
               // validColumn[1] = validColumn[4] = validColumn[5] = validColumn[9] = validColumn[10] = validColumn[25] = validColumn[26] = validColumn[27] = validColumn[28] = validColumn[30] = validColumn[31] = validColumn[45] = validColumn[46] = validColumn[48] = validColumn[49] = 1;
                // d4 e5 i9 j10 y25 z26 aa27 ab28 ad30 ae31 as45 at46 aw48 ax49
            }
            else
                row = 0;   
        }

        private void BrowseButton1_MouseDown(object sender, MouseButtonEventArgs e)
        {
            browsefile1.Margin = new Thickness(0, 0, 2, 0);
        }

        private void BrowseButton2_MouseDown(object sender, MouseButtonEventArgs e)
        {
            browsefile2.Margin = new Thickness(0, 0, 2, 0);
        }

        private void CompareButton_MouseDown(object sender, MouseButtonEventArgs e)
        {
            CompareButton.Margin = new Thickness(2);
        }

        private void close_MouseDown(object sender, MouseButtonEventArgs e)
        {
            close.Margin = new Thickness(2);
        }

        void showanimation_Completed(object sender, EventArgs e)
        {
            DoubleAnimation showanimation = new DoubleAnimation(1d, 1d, TimeSpan.FromMilliseconds(5000));
            showanimation.Completed += showanimation1_Completed;
            error_msg_border.BeginAnimation(OpacityProperty, showanimation);

        }

        private void showanimation1_Completed(object sender, EventArgs e)
        {
            DoubleAnimation hideAnimation = new DoubleAnimation(1d, 0d, TimeSpan.FromMilliseconds(500));
            error_msg_border.BeginAnimation(OpacityProperty, hideAnimation); 
        }
       
        private void close_MouseUp(object sender, MouseButtonEventArgs e)
        {
            close.Margin = new Thickness(0);
            Application.Current.Shutdown();
        }

        private void minimize_MouseUp(object sender, MouseButtonEventArgs e)
        {
            WindowState = WindowState.Minimized;
        }

        private void Grid_MouseDown(object sender, MouseButtonEventArgs e)
        {
            DragMove();
        }
     
    }
}
