#region Imports

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

#endregion

namespace VBTranspiler.CodeGenerator.Controls
{
  public abstract class ControlPropertyBase
  {
    private string mPropertyName;

    public ControlPropertyBase(string propertyName)
    {
      mPropertyName = propertyName;
    }
  }

  public class SimpleControlProperty : ControlPropertyBase
  {
    private string mValue;

    public SimpleControlProperty(string propertyName, string value)
      : base(propertyName)
    {
      mValue = value;
    }
  }

  public class NestedControlProperty : ControlPropertyBase
  {
    private int mArrayIndex;
    private List<ControlPropertyBase> mProperties;

    public NestedControlProperty(string propertyName, int arrayIndex, IEnumerable<ControlPropertyBase> properties)
      : base(propertyName)
    {
      mArrayIndex = arrayIndex;
      mProperties = new List<ControlPropertyBase>(properties);
    }
  }
}
