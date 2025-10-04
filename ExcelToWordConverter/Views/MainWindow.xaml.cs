using ExcelToWordConverter.ViewModels;
using System;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Controls;

namespace ExcelToWordConverter.Views
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
            DataContext = new MainViewModel();

            // Явная подписка на событие нажатия кнопки с обработкой уже обработанных событий
            this.Loaded += (s, e) =>
            {
                var button = FindName("ConvertButton") as Button; // Укажите x:Name вашей кнопки в XAML
                if (button != null)
                {
                    button.AddHandler(Button.ClickEvent, new RoutedEventHandler(OnConvertButtonClick), true);
                }
            };
        }

        private async void OnConvertButtonClick(object sender, RoutedEventArgs e)
        {
            if (DataContext is MainViewModel viewModel)
            {
                if (viewModel.ConvertCommand.CanExecute(null))
                {
                    await viewModel.ConvertCommand.ExecuteAsync(null);
                }
            }
        }
    }

}
