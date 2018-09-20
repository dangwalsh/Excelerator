namespace Gensler.Revit.Excelerator.Views
{
    using System;
    using System.Windows.Forms;
    using System.Windows.Input;

    class FileCommand : ICommand
    {
        private readonly MainWindowViewModel _viewModel;

        public event EventHandler CanExecuteChanged;

        public bool CanExecute(object parameter)
        {
            return true;
        }

        public void Execute(object parameter)
        {
            var openFileDialog = new OpenFileDialog();
            if (openFileDialog.ShowDialog() == DialogResult.OK)
                _viewModel.ExcelPath = openFileDialog.FileName;
        }

        public FileCommand(MainWindowViewModel viewModel)
        {
            _viewModel = viewModel;
        }
    }
}
