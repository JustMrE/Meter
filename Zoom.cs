
namespace Meter
{
    class Zoom
{
     public Zoom(object InitialValue)
     { 
          _value = InitialValue;
     }

     public event EventHandler<ValueChangedEventArgs> ValueChanged;

     protected virtual void OnValueChanged(ValueChangedEventArgs e)
     {
           if(ValueChanged != null)
              ValueChanged(this, e);
     }

     private object _value;

     public object Value
     {
         get { return _value; }
         set 
         {
             object oldValue = _value;
             _value = value;
             OnValueChanged(new ValueChangedEventArgs(oldValue, _value));
         }
     }
}
}