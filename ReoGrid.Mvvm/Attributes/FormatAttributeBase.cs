using ReoGrid.Mvvm.Interfaces;
using System;
using unvell.ReoGrid.DataFormat;

namespace ReoGrid.Mvvm.Attributes
{
    public abstract class FormatAttributeBase : Attribute, IFormatArgs
    {
        public abstract CellDataFormatFlag CellDataFormatFlag { get; }
    }
}
