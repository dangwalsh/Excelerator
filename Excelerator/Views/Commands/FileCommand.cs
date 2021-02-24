namespace Gensler.Revit.Excelerator.Views
{
    using System;
    using System.Windows.Forms;
    using System.Windows.Input;

    class FileCommand : ICommand
    {
        private readonly MainWindowViewModel _viewModel;

        public event EventHandler CanExecuteChanged
        {
            add { CommandManager.RequerySuggested += value; }
            remove { CommandManager.RequerySuggested -= value; }
        }

        public bool CanExecute(object parameter)
        {
            return true;
        }

        public void Execute(object parameter)
        {
            var openFileDialog = new OpenFileDialog();
            if (openFileDialog.ShowDialog() != DialogResult.OK) return;

            _viewModel.ExcelPath = openFileDialog.FileName;
            _viewModel.IsLoaded = true;
        }

        public FileCommand(MainWindowViewModel viewModel)
        {
            _viewModel = viewModel;
        }
    }
}
