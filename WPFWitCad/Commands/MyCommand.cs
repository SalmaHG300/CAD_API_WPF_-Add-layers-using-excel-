using System;
using System.Windows.Input;

namespace WPFWitCad.Commands
{
  public class MyCommand : ICommand
  {
    public event EventHandler CanExecuteChanged;

    public Action<Object> ExcuteCmd { get; set; }

    public Predicate<object> CanExcuteCmd { get; set; }


    #region Constructor
    public MyCommand(Action<Object> _excute, Predicate<object> _canExcute)
    {
      ExcuteCmd = _excute;

      CanExcuteCmd = _canExcute;
    }
    #endregion

    #region Methods


    public bool CanExecute(object parameter)
    {
      return true;
    }

    public void Execute(object parameter)
    {
      ExcuteCmd(parameter);
    }
    #endregion


  }
}
