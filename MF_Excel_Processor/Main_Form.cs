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
        ExcelData input;

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
        /// Enables the lopp worker thread.
        /// </summary>
        public bool WorkerThreadEnabled { 
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
        /// Stores the input file location.
        /// </summary>
        private string fileName;

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
                Marshal.ReleaseComObject(input.fullRange);
                Marshal.ReleaseComObject(input.currentSheet);
                input.currentWorkbook.Close();
                Marshal.ReleaseComObject(input.currentWorkbook);
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
                //Conditional variables or counters.
                int nullCount = 0;
                int maxRows = input.fullRange.Rows.Count;
                int count = 0;

                //Input file traversing variables.
                int currentPosition = startingRows[0];
                int columnA = 1;
                string typeA = input.fullRange.Cells[startingRows[1], columnA + 1].Value2;
                string typeB = input.fullRange.Cells[startingRows[1] + 1, columnA + 1].Value2;
                string typeC = input.fullRange.Cells[startingRows[1], columnA + 3].Value2;
                string typeD = input.fullRange.Cells[startingRows[1] + 1, columnA + 3].Value2;

                

                // Iteration through cells.
                while (nullCount < 3 && workerThreadEnabled)
                {
                    //If not null
                    if (input.fullRange.Cells[currentPosition, columnA] != null && input.fullRange.Cells[currentPosition, columnA].Value2 != null)
                    {
                        if(input.fullRange.Cells[currentPosition, columnA].Value2 is double number)
                        {
                            //Copy data values;
                            output.currentSheet.Cells[newSheetPositionY, newSheetPositionX] = number;
                            output.currentSheet.Cells[newSheetPositionY, newSheetPositionX + 1] = input.fullRange.Cells[currentPosition, columnA + 1].Value2;
                            output.currentSheet.Cells[newSheetPositionY, newSheetPositionX + 2] = input.fullRange.Cells[currentPosition, columnA + 2].Value2;
                            output.currentSheet.Cells[newSheetPositionY, newSheetPositionX + 3] = input.fullRange.Cells[currentPosition, columnA + 3].Value2;

                            //Copy type values.
                            output.currentSheet.Cells[newSheetPositionY, newSheetPositionX + 4] = typeA;
                            output.currentSheet.Cells[newSheetPositionY, newSheetPositionX + 5] = typeB;
                            output.currentSheet.Cells[newSheetPositionY, newSheetPositionX + 6] = typeC;
                            output.currentSheet.Cells[newSheetPositionY, newSheetPositionX + 7] = typeD;

                            newSheetPositionY++;
                        }
                        else
                        {
                            //New type section hit, save values.
                            typeA = input.fullRange.Cells[currentPosition, columnA + 1].Value2;
                            typeB = input.fullRange.Cells[currentPosition + 1, columnA + 1].Value2;
                            typeC = input.fullRange.Cells[currentPosition, columnA + 3].Value2;
                            typeD = input.fullRange.Cells[currentPosition + 1, columnA + 3].Value2;
                            currentPosition += 2;
                        }
                    }
                    else nullCount++;
                    currentPosition++;
                    count++;

                    //Used to report progress.
                    if (count > 5)
                    {
                        backgroundWorker2.ReportProgress(100 * currentPosition / maxRows);
                        count = 0;
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
                input = new ExcelData(fileName);
                IsFileLoaded = true;
            }
        }

        /// <summary>
        /// Background worker 2 work method.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void BackgroundWorker2_DoWork(object sender, DoWorkEventArgs e)
        {
            if(state == State.Working)
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
                    input.dataColumnsReady = false;
                    input.typeColumnsReady = false;

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
                if (openDialog.ShowDialog() == DialogResult.OK)
                {
                    if (IsFileLoaded) Cleanup();
                    state = State.LoadingFile;
                    fileName = openDialog.FileName;
                    OpenFileButton.Enabled = false;
                    DataTextBox.Text = "Archivo cargado: \n\n" + openDialog.FileName;
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
                input.dataColumnsReady = true;
                RowBox1.BackColor = Color.LightGreen;
                input.typeColumnsReady = true;
                RowBox2.BackColor = Color.LightGreen;
                MainTextBox.Text = input.fullRange.Rows.Count.ToString();
            }
            catch (Exception ex)
            {
                input.dataColumnsReady = false;
                input.typeColumnsReady = false;
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
            if (input.dataColumnsReady && input.typeColumnsReady)
            {
                state = State.Working;
                StartButton.Enabled = false;
                //Create new instance.
                output = new ExcelData(null);
                output.excelApp.Visible = false;

                timer.Start();
                activeThreads = 1;
                workerThreadEnabled = true;
                backgroundWorker2.RunWorkerAsync();
            }
        }
    }
}
