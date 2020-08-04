/*
 * MF_Excel_Processor
 * Main_Form.cs
 * Author: Cesar Zavala
 * 
 */
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
using System.Runtime.InteropServices;
using System.Reflection;
using System.IO;
using System.Security.Authentication;
using Microsoft.CSharp.RuntimeBinder;
using Microsoft.Office.Interop.Excel;
using System.Collections.Concurrent;
using System.Windows.Forms.VisualStyles;
using System.Diagnostics;
using System.CodeDom;

namespace MF_Excel_Processor
{
    public partial class Main_Form : Form
    {
        /// <summary>
        /// Input file to be processed.
        /// </summary>
        ExcelData[] inputFiles;

        /// <summary>
        /// Output file of processed data.
        /// </summary>
        ExcelData output;

        /// <summary>
        /// First background worker.
        /// </summary>
        private BackgroundWorker backgroundWorker1 = new BackgroundWorker();

        /// <summary>
        /// Second background worker.
        /// </summary>
        private BackgroundWorker backgroundWorker2 = new BackgroundWorker();

        /// <summary>
        /// Backing variable for WorkerThreadEnabled.
        /// </summary>
        private bool workerThreadEnabled = false;

        /// <summary>
        /// Input file format.
        /// </summary>
        private InputType type;

        /// <summary>
        /// Enables the lopp worker thread.
        /// </summary>
        public bool WorkerThreadEnabled
        {
            get
            {
                return workerThreadEnabled;
            }
            set
            {
                workerThreadEnabled = value;
                if (workerThreadEnabled)
                {
                    CancelButton.Enabled = true;
                    StartButton.Enabled = false;
                    OpenFileButton.Enabled = false;
                    RowConfirmButton.Enabled = false;
                }
                else
                {
                    CancelButton.Enabled = false;
                    OpenFileButton.Enabled = true;
                    StartButton.Enabled = true;
                    RowConfirmButton.Enabled = true;
                }
            }
        }

        /// <summary>
        /// Rows where the data starts.
        /// </summary>
        private int[] startingRows;

        /// <summary>
        /// Keeps track of the state of the program.
        /// </summary>
        private State state = State.Idle;

        /// <summary>
        /// Keeps track of all active threads.
        /// </summary>
        int activeThreads;

        /// <summary>
        /// Determines if a file is loaded.
        /// </summary>
        public bool IsFileLoaded { get; private set; } = false;

        /// <summary>
        /// Stores the input files' location.
        /// </summary>
        private string[] fileNames;

        /// <summary>
        /// Keeps track of the elapsed time.
        /// </summary>
        private Stopwatch timer = new Stopwatch();

        /// <summary>
        /// States of the program.
        /// </summary>
        public enum State
        {
            Idle,
            Working,
            LoadingFile
        }

        /// <summary>
        /// Form constructor.
        /// </summary>
        public Main_Form()
        {
            InitializeComponent();

            //Fill up SelectionBox
            SelectionBox.Items.AddRange(Enum.GetNames(typeof(InputType)));
            type = InputType.Agricola;
            SelectionBox.SelectedItem = type.ToString();

            //Event Handlers
            this.FormClosing += new FormClosingEventHandler(Closing);

            //Worker setup.
            backgroundWorker1.DoWork += BackgroundWorker1_DoWork;
            backgroundWorker2.DoWork += BackgroundWorker2_DoWork;
            backgroundWorker1.RunWorkerCompleted += BackgroundWorkers_RunWorkerCompleted;
            backgroundWorker2.RunWorkerCompleted += BackgroundWorkers_RunWorkerCompleted;
            backgroundWorker2.ProgressChanged += Worker2_ProgressChanged;
            backgroundWorker2.WorkerReportsProgress = true;

        }

        ///////////////////////////////////////////////
        ///Memory and closing functions.
        ///////////////////////////////////////////////

        /// <summary>
        /// Releases input file and calls garbage collector.
        /// </summary>
        private void Cleanup()
        {
            StartButton.Enabled = false;
            try
            {
                //Garbage collector.
                GC.Collect();
                GC.WaitForPendingFinalizers();

                //Original file cleanup
                foreach (ExcelData input in inputFiles)
                {
                    Marshal.ReleaseComObject(input.fullRange);
                    Marshal.ReleaseComObject(input.currentSheet);
                    input.currentWorkbook.Close();
                    Marshal.ReleaseComObject(input.currentWorkbook);
                }
                DataTextBox.Clear();
                IsFileLoaded = false;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: No hay documento para cerrar.");
            }
        }

        /// <summary>
        /// Actions to be carried out when the program is closing.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            if (IsFileLoaded) Cleanup();
        }

        ///////////////////////////////////////////////
        ///Processing functions.
        ///////////////////////////////////////////////

        /// <summary>
        /// Transfers a specified column on the Excel file into a new one.
        /// </summary>
        /// <param name="newSheetPositionX">Starting column cell in which the parsed data is copied.</param>
        /// <param name="newSheetPositionY">Starting row cell in which the parsed data is copied.</param>
        private void startProcessing(int newSheetPositionX, int newSheetPositionY)
        {
            try
            {
                foreach (ExcelData input in inputFiles)
                {
                    //Conditional variables or counters.
                    int nullCount = 0;
                    int maxRows = input.fullRange.Rows.Count;
                    int count = 0;

                    //Input file traversing variables.
                    int currentRow = startingRows[0];
                    int columnA = 1;
                    int currentColumn = newSheetPositionX;
                    string typeA = input.fullRange.Cells[startingRows[1], columnA + 1].Value2;
                    string typeB = input.fullRange.Cells[startingRows[1] + 1, columnA + 1].Value2;
                    string typeC = input.fullRange.Cells[startingRows[1], columnA + 3].Value2;
                    string typeD = input.fullRange.Cells[startingRows[1] + 1, columnA + 3].Value2;

                    //Enum dependent column setup
                    int maxColumn;
                    switch (type)
                    {
                        case InputType.Agricola:
                            maxColumn = newSheetPositionX + 4;
                            break;
                        case InputType.Pecuario:
                            maxColumn = newSheetPositionX + 5;
                            break;
                        default:
                            throw new Exception("Invalid Input type");
                            break;
                    }


                    // Iteration through cells.
                    while (nullCount < 3 && workerThreadEnabled)
                    {
                        //If not null
                        if (input.fullRange.Cells[currentRow, columnA] != null && input.fullRange.Cells[currentRow, columnA].Value2 != null)
                        {
                            if (input.fullRange.Cells[currentRow, columnA].Value2 is double number)
                            {
                                //Copy data values;
                                output.currentSheet.Cells[newSheetPositionY, currentColumn++] = number;
                                int i = 1;
                                while(currentColumn < maxColumn) 
                                {
                                    output.currentSheet.Cells[newSheetPositionY, currentColumn] = input.fullRange.Cells[currentRow, columnA + i].Value2;
                                    i++;
                                    currentColumn++;
                                }

                                //Copy type values.
                                output.currentSheet.Cells[newSheetPositionY, currentColumn++] = typeA;
                                output.currentSheet.Cells[newSheetPositionY, currentColumn++] = typeB;
                                output.currentSheet.Cells[newSheetPositionY, currentColumn++] = typeC;
                                output.currentSheet.Cells[newSheetPositionY, currentColumn++] = typeD;
                                currentColumn = newSheetPositionX;

                                newSheetPositionY++;
                            }
                            else
                            {
                                //New type section hit, save values.
                                typeA = input.fullRange.Cells[currentRow, columnA + 1].Value2;
                                typeB = input.fullRange.Cells[currentRow + 1, columnA + 1].Value2;
                                typeC = input.fullRange.Cells[currentRow, columnA + 3].Value2;
                                typeD = input.fullRange.Cells[currentRow + 1, columnA + 3].Value2;
                                currentRow += 2;
                            }
                        }
                        else nullCount++;
                        currentRow++;
                        count++;

                        //Used to report progress.
                        if (count > 5)
                        {
                            backgroundWorker2.ReportProgress(100 * currentRow / maxRows);
                            count = 0;
                        }
                    }
                }
            }

            //Error handling
            catch (Exception ex)
            {
                String errorMessage;
                errorMessage = "Error: ";
                errorMessage = String.Concat(errorMessage, ex.Message);
                errorMessage = String.Concat(errorMessage, "\n Full String: ");
                errorMessage = String.Concat(errorMessage, ex.ToString());
                MessageBox.Show(errorMessage, "Error");
            }
        }

        ///////////////////////////////////////////////
        ///Background worker functions.
        ///////////////////////////////////////////////

        /// <summary>
        /// Background worker 1 work method.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void BackgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
        {
            if (state == State.LoadingFile)
            {
                inputFiles = new ExcelData[fileNames.Length];
                for (int i = 0; i < fileNames.Length; i++)
                {
                    inputFiles[i] = new ExcelData(fileNames[i],type);
                }
            }
        }

        /// <summary>
        /// Background worker 2 work method.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void BackgroundWorker2_DoWork(object sender, DoWorkEventArgs e)
        {
            if (state == State.Working)
            {
                startProcessing(1, 3);
            }
            workerThreadEnabled = false;
        }

        /// <summary>
        /// Updates the progress text.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Worker2_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            MainTextBox.Text = "Progreso: %" + e.ProgressPercentage;
        }

        /// <summary>
        /// Work completed event handler
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void BackgroundWorkers_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if (state == State.LoadingFile)
            {
                IsFileLoaded = true;
                StartButton.Enabled = true;
                RowConfirmButton.Enabled = true;
                OpenFileButton.Enabled = true;
                state = State.Idle;
            }

            else
            {
                activeThreads--;

                //If no more threads are running.
                if (activeThreads < 1)
                {
                    WorkerThreadEnabled = false;
                    inputFiles[0].dataColumnsReady = false;
                    inputFiles[0].typeColumnsReady = false;

                    //Displays the total time it took to carry out the processing.
                    if (timer.IsRunning)
                    {
                        timer.Stop();
                        TimeSpan ts = timer.Elapsed;
                        string totalTime = String.Format("{0:00}:{1:00}:{2:00}.{3:00}", ts.Hours, ts.Minutes, ts.Seconds, ts.Milliseconds / 10);
                        MainTextBox.Text = "Tiempo total: " + totalTime;
                    }
                    RowBox1.BackColor = Color.White;
                    RowBox2.BackColor = Color.White;
                    output.excelApp.Visible = true;

                    //Reset counters.
                    timer.Reset();

                    state = State.Idle;
                }
            }
        }

        ///////////////////////////////////////////////
        ///Button event handlers.
        ///////////////////////////////////////////////

        /// <summary>
        /// "Open File" button event handler.
        /// </summary>
        /// <param name="sender">The button.</param>
        /// <param name="e">Event.</param>
        private void OpenFileButton_Click(object sender, EventArgs e)
        {
            //File input.
            try
            {
                OpenFileDialog openDialog = new OpenFileDialog();
                openDialog.Multiselect = true;
                if (openDialog.ShowDialog() == DialogResult.OK)
                {
                    if (IsFileLoaded) Cleanup();
                    state = State.LoadingFile;
                    fileNames = openDialog.FileNames;
                    OpenFileButton.Enabled = false;
                    DataTextBox.Text = "Archivos cargados: \n\n" + openDialog.FileNames.Length;
                    type = (InputType)Enum.Parse(typeof(InputType), SelectionBox.SelectedItem.ToString());
                    backgroundWorker1.RunWorkerAsync();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error cargando el archivo, verifique el formato.");
            }
        }

        /// <summary>
        /// Cleanup button event handler.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void CleanupButton_Click(object sender, EventArgs e)
        {
            Cleanup();
        }

        /// <summary>
        /// Confirms rows required to obtain data types and data values.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void RowConfirmButton_Click(object sender, EventArgs e)
        {
            try
            {
                startingRows = new int[2];
                startingRows[0] = Int32.Parse(RowBox1.Text);
                startingRows[1] = Int32.Parse(RowBox2.Text);
                inputFiles[0].dataColumnsReady = true;
                RowBox1.BackColor = Color.LightGreen;
                inputFiles[0].typeColumnsReady = true;
                RowBox2.BackColor = Color.LightGreen;
            }
            catch (Exception ex)
            {
                inputFiles[0].dataColumnsReady = false;
                inputFiles[0].typeColumnsReady = false;
                RowBox1.BackColor = Color.White;
                RowBox2.BackColor = Color.White;
                MessageBox.Show("Error confirmando filas.");
            }
        }

        /// <summary>
        /// Cancels active threads.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void CancelButton_Click(object sender, EventArgs e)
        {
            WorkerThreadEnabled = false;
        }

        /// <summary>
        /// Start button event handler.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void StartButton_Click(object sender, EventArgs e)
        {
            if (inputFiles[0].dataColumnsReady && inputFiles[0].typeColumnsReady)
            {
                state = State.Working;
                //Create new instance.
                output = new ExcelData(null,type);
                output.excelApp.Visible = false;

                timer.Start();
                activeThreads = 1;
                workerThreadEnabled = true;
                backgroundWorker2.RunWorkerAsync();
            }
        }
    }
}
