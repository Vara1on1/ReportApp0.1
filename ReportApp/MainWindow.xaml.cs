﻿using ReportApp.Model;
using System;
using System.Collections.Generic;
using System.Data.Entity;
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
using Excel = Microsoft.Office.Interop.Excel;
using ReportApp.View;
using ReportApp.View.Pages;

namespace ReportApp
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        

        public MainWindow()
        {
            InitializeComponent();
            MainFrame.Navigate(new AuthoregPage());
        }

        

        private void UpdateChart(object sender, SelectionChangedEventArgs e)
        {

        }

        private void UpdateCharter(object sender, SelectionChangedEventArgs e)
        {

        }

        private void BtnExportToExcel_Click(object sender, RoutedEventArgs e)
        {

        }
    }
}
