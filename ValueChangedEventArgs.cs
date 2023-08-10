

namespace Meter
{
    class ValueChangedEventArgs : EventArgs
{
    public readonly object LastValue;
    public readonly object NewValue;

    public ValueChangedEventArgs(object LastValue, object NewValue)
    {
        this.LastValue = LastValue;
        this.NewValue = NewValue;
    }
}
}