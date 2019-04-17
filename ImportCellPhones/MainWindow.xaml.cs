/* Title:           Cell Phone Imports
 * Date:            4-17-19
 * Author:          Terry Holmes
 * 
 * Description:     This application is used to import the cell phones */

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
using Excel = Microsoft.Office.Interop.Excel;
using DataValidationDLL;
using PhonesDLL;
using NewEventLogDLL;
using NewEmployeeDLL;

namespace ImportCellPhones
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        //setting up the classes
        DataValidationClass TheDataValidationClass = new DataValidationClass();
        WPFMessagesClass TheMessagesClass = new WPFMessagesClass();
        PhonesClass ThePhoneClass = new PhonesClass();
        EventLogClass TheEventLogClass = new EventLogClass();
        EmployeeClass TheEmployeeClass = new EmployeeClass();

        //setting up the data
        ImportPhonesDataSet TheImportPhonesDataSet = new ImportPhonesDataSet();
        FindEmployeeByLastNameDataSet TheFindEmployeeByLastNameDataSet = new FindEmployeeByLastNameDataSet();
        FindWarehousesDataSet TheFindWarehouseDataSet = new FindWarehousesDataSet();

        public MainWindow()
        {
            InitializeComponent();
        }

        private void BtnImportExcel_Click(object sender, RoutedEventArgs e)
        {
            Excel.Application xlDropOrder;
            Excel.Workbook xlDropBook;
            Excel.Worksheet xlDropSheet;
            Excel.Range range;

            int intColumnRange = 0;
            int intCounter;
            int intNumberOfRecords;
            string strFirstName;
            string strLastName;
            string strCellNumber;
            int intRecordsReturned;
            int intEmployeeCounter;
            int intEmployeeID;
            int intWarehouseID = 0;
            string strPhoneNotes = "PHONE IMPORTED";

            try
            {
                TheImportPhonesDataSet.cellphones.Rows.Clear();

                Microsoft.Win32.OpenFileDialog dlg = new Microsoft.Win32.OpenFileDialog();
                dlg.FileName = "Document"; // Default file name
                dlg.DefaultExt = ".xlsx"; // Default file extension
                dlg.Filter = "Excel (.xlsx)|*.xlsx"; // Filter files by extension

                // Show open file dialog box
                Nullable<bool> result = dlg.ShowDialog();

                // Process open file dialog box results
                if (result == true)
                {
                    // Open document
                    string filename = dlg.FileName;
                }

                PleaseWait PleaseWait = new PleaseWait();
                PleaseWait.Show();

                xlDropOrder = new Excel.Application();
                xlDropBook = xlDropOrder.Workbooks.Open(dlg.FileName, 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                xlDropSheet = (Excel.Worksheet)xlDropOrder.Worksheets.get_Item(1);

                range = xlDropSheet.UsedRange;
                intNumberOfRecords = range.Rows.Count;
                intColumnRange = range.Columns.Count;

                for (intCounter = 1; intCounter <= intNumberOfRecords; intCounter++)
                {
                    strFirstName = Convert.ToString((range.Cells[intCounter, 3] as Excel.Range).Value2).ToUpper();
                    strLastName = Convert.ToString((range.Cells[intCounter, 4] as Excel.Range).Value2).ToUpper();
                    strCellNumber = Convert.ToString((range.Cells[intCounter, 2] as Excel.Range).Value2).ToUpper();
                                       
                    TheFindEmployeeByLastNameDataSet = TheEmployeeClass.FindEmployeesByLastNameKeyWord(strLastName);

                    intRecordsReturned = TheFindEmployeeByLastNameDataSet.FindEmployeeByLastName.Rows.Count - 1;
                    intEmployeeID = -1;

                    if (intRecordsReturned > -1)
                    {
                        for (intEmployeeCounter = 0; intEmployeeCounter <= intRecordsReturned; intEmployeeCounter++)
                        {
                            if (strFirstName == TheFindEmployeeByLastNameDataSet.FindEmployeeByLastName[intEmployeeCounter].FirstName)
                            {
                                intEmployeeID = TheFindEmployeeByLastNameDataSet.FindEmployeeByLastName[intEmployeeCounter].EmployeeID;
                                intWarehouseID = FindWarehouseID(TheFindEmployeeByLastNameDataSet.FindEmployeeByLastName[intEmployeeCounter].HomeOffice);
                            }
                        }
                    }

                    ImportPhonesDataSet.cellphonesRow NewPhoneRow = TheImportPhonesDataSet.cellphones.NewcellphonesRow();

                    NewPhoneRow.EmployeeID = intEmployeeID;
                    NewPhoneRow.FirstName = strFirstName;
                    NewPhoneRow.LastName = strLastName;
                    NewPhoneRow.CellNumber = strCellNumber;
                    NewPhoneRow.WarehouseID = intWarehouseID;
                    NewPhoneRow.PhoneNotes = strPhoneNotes;

                    TheImportPhonesDataSet.cellphones.Rows.Add(NewPhoneRow);
                }

                PleaseWait.Close();
                dgrResults.ItemsSource = TheImportPhonesDataSet.cellphones;
            }
            catch (Exception Ex)
            {
                TheEventLogClass.InsertEventLogEntry(DateTime.Now, "Import Cell Phones // Import Excel Button " + Ex.Message);

                TheMessagesClass.ErrorMessage(Ex.ToString());
            }
           
        }
        private int FindWarehouseID(string strWarehouse)
        {
            int intWarehouseID = 0;
            int intCounter;
            int intNumberOfRecords;

            TheFindWarehouseDataSet = TheEmployeeClass.FindWarehouses();

            intNumberOfRecords = TheFindWarehouseDataSet.FindWarehouses.Rows.Count - 1;

            for(intCounter = 0; intCounter <= intNumberOfRecords; intCounter++)
            {
                if(strWarehouse == TheFindWarehouseDataSet.FindWarehouses[intCounter].FirstName)
                {
                    intWarehouseID = TheFindWarehouseDataSet.FindWarehouses[intCounter].EmployeeID;
                }
            }

            return intWarehouseID;
        }

        private void BtnProcess_Click(object sender, RoutedEventArgs e)
        {
            int intCounter;
            int intNumberOfRecords;
            bool blnFatalError = false;
            string strPhoneNumber;
            int intEmployeeID;
            int intWarehouseID;
            string strPhoneNotes;

            try
            {
                intNumberOfRecords = TheImportPhonesDataSet.cellphones.Rows.Count - 1;

                for(intCounter = 0; intCounter <= intNumberOfRecords; intCounter++)
                {
                    strPhoneNotes = TheImportPhonesDataSet.cellphones[intCounter].PhoneNotes;
                    strPhoneNumber = TheImportPhonesDataSet.cellphones[intCounter].CellNumber;
                    intEmployeeID = TheImportPhonesDataSet.cellphones[intCounter].EmployeeID;
                    intWarehouseID = TheImportPhonesDataSet.cellphones[intCounter].WarehouseID;

                    
                }
            }
            catch (Exception Ex)
            {
                TheEventLogClass.InsertEventLogEntry(DateTime.Now, "Import Cell Phones // Main Window // Process Button " + Ex.Message);

                TheMessagesClass.ErrorMessage
            }
        }
    }
}
