using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;
using WPFWitCad.Commands;
using WPFWitCad.DataContext;
using WPFWitCad.Model;

namespace WPFWitCad.ViewModel
{
  public class MainWindowViewModel :INotifyPropertyChanged
  {
    #region Constructor
    public MainWindowViewModel()

    {
      CadLayers = AutocadData.GetCadLayers();

      CreateBtnCmd = new MyCommand(CreateExcuteCmd, CanCreateExcuteCmd);

      //UpdateBtnCmd = new MyCommand(UpdateExcuteCmd, CanUpdateExcuteCmd);


    }

        //private bool CanUpdateExcuteCmd(object obj)
        //{
        //    return true;
        //}

        //private void UpdateExcuteCmd(object obj)
        //{            
        //    // After updating the layers, you may need to update the CadLayers collection
        //    CadLayers = AutocadData.GetCadLayers();
        //    OnproperyChanged(nameof(CadLayers));
        //}
        #endregion

        #region Properties

        public List<CadLayerObj> CadLayers { get; set; } = new List<CadLayerObj>();

    private CadLayerObj _selectedLayer;

    public event PropertyChangedEventHandler PropertyChanged;

    public CadLayerObj SelectedLayer
    {
      get { return _selectedLayer; }
      set { _selectedLayer = value;
           OnproperyChanged();


      }
    }

    public MyCommand CreateBtnCmd { get; set; }

     //public MyCommand UpdateBtnCmd { get; set; }



        #endregion

        #region Methods
        public void CreateExcuteCmd(object parameter)
        {

             AutocadData.Getkeywords(SelectedLayer.Name);
            // Notify the UI that the CadLayers collection has changed         

        }

    public bool CanCreateExcuteCmd(object parameter)
    
    {
      return true;
    
    }


    public void OnproperyChanged([CallerMemberName] string Name=null)
    {

      PropertyChanged.Invoke(this, new PropertyChangedEventArgs(Name));

    }
    #endregion

  }
}
