using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using ExcelToWordConverter.Models;
using Microsoft.Win32;
using System;
using System.Collections.ObjectModel;
using System.IO;
using System.Threading.Tasks;
using System.Windows;

namespace ExcelToWordConverter.ViewModels
{

    public partial class MainViewModel : ObservableObject
    {
        [ObservableProperty]
        private ObservableCollection<string> files = new();

        [ObservableProperty]
        private int currentYear = DateTime.Now.Year;

        [ObservableProperty]
        private string status = "Ожидание действий...";

        [ObservableProperty]
        private bool isBusy;

        [RelayCommand]
        private void SelectFiles()
        {
            var dlg = new OpenFileDialog
            {
                Multiselect = true,
                Filter = "Excel files (*.xlsx)|*.xlsx"
            };
            if (dlg.ShowDialog() == true)
            {
                Files.Clear();
                foreach (var file in dlg.FileNames)
                    Files.Add(file);
            }
        }

        [RelayCommand]
        private void ClearFiles()
        {
            Files.Clear();
            Status = "Список очищен.";
        }


        [RelayCommand]
        private async Task ConvertAsync()
        {
            if (Files.Count == 0)
            {
                MessageBox.Show("Выберите файлы Excel.");
                return;
            }

            IsBusy = true;
            Status = "Конвертация...";

            try
            {
                foreach (var file in Files)
                {
                    string output = Path.ChangeExtension(file, ".docx");
                    await ExamConverter.ConvertAsync(file, output);
                }
                Status = "Готово!";
                MessageBox.Show("Конвертация завершена успешно.");
            }
            catch (Exception ex)
            {
                Status = "Ошибка.";
                MessageBox.Show(ex.Message, "Ошибка");
            }
            finally
            {
                IsBusy = false;
            }
        }
    }
}
