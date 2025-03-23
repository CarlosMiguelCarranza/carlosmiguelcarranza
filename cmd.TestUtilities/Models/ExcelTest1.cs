using cmc.ExcelUtilities;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace cmd.TestUtilities.Models
{
    public class ExcelTest1 : ExcelSheetSchemaBase
    {
        [ExcelColumn("A")]
        public bool BoolProperty { get; set; }

        [ExcelColumn("B")]
        public char CharProperty { get; set; }

        [ExcelColumn("C")]
        public int IntProperty { get; set; }

        [ExcelColumn("D")]
        public decimal DecimalProperty { get; set; }

        [ExcelColumn("E")]
        public double DoubleProperty { get; set; }

        [ExcelColumn("F")]
        public float FloatProperty { get; set; }

        [ExcelColumn("G")]
        public long LongProperty { get; set; }

        [ExcelColumn("H")]
        public DateTime DateTimeProperty { get; set; }

        [ExcelColumn("I")]
        public Guid GuidProperty { get; set; }

        [ExcelColumn("J")]
        public TimeSpan TimeSpanProperty { get; set; }

        // Tipos nullables
        [ExcelColumn("K")]
        public bool? NullableBoolProperty { get; set; }

        [ExcelColumn("L")]
        public char? NullableCharProperty { get; set; }

        [ExcelColumn("M")]
        public int? NullableIntProperty { get; set; }

        [ExcelColumn("N")]
        public decimal? NullableDecimalProperty { get; set; }

        [ExcelColumn("O")]
        public double? NullableDoubleProperty { get; set; }

        [ExcelColumn("Q")]
        public float? NullableFloatProperty { get; set; }

        [ExcelColumn("R")]
        public long? NullableLongProperty { get; set; }

        [ExcelColumn("S")]
        public DateTime? NullableDateTimeProperty { get; set; }

        [ExcelColumn("T")]
        public Guid? NullableGuidProperty { get; set; }

        [ExcelColumn("U")]
        public TimeSpan? NullableTimeSpanProperty { get; set; }

        // Tipos adicionales
        [ExcelColumn("V")]
        public byte ByteProperty { get; set; }

        [ExcelColumn("W")]
        public byte? NullableByteProperty { get; set; }

        [ExcelColumn("X")]
        public string StringProperty { get; set; } // string es nullable por defecto

        [ExcelColumn("Y")]
        public byte[] ByteArrayProperty { get; set; }
    }
}
