using Autodesk.AutoCAD.Colors;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Media;
using Color = Autodesk.AutoCAD.Colors.Color;

namespace WPFWitCad.Model
{
  public class CadLayerObj
  {
        #region MyRegion
        public CadLayerObj()
        {
              
        }
    #endregion

    #region properties
    public string Name { get; set; }

    public Color Color { get; set; }

    public string LayerLineTtpe { get; set; }  


    #endregion
    #region Method
    public override string ToString()
    {
      return Name;
    }
    #endregion
  }
}
