using System;
using System.Collections.Generic;

namespace ExcelSharp.NPOI
{
    public class SheetRangeRowSelector
    {
        private SheetRange Range;

        public SheetRangeRowSelector(SheetRange range)
        {
            Range = range;
        }

        public SheetRange this[int? startOffset, int? endOffset]
        {
            get
            {
                if (startOffset is null) startOffset = 0;
                if (endOffset is null) endOffset = Range.RowLength - 1;

                return new SheetRange(Range.Sheet, (Range.Start.Row + startOffset.Value, Range.Start.Col), (Range.Start.Row + endOffset.Value, Range.End.Col));
            }
        }

        public SheetRangeGroup Select(Func<int, bool> indexSelector)
        {
            IEnumerable<SheetRange> select()
            {
                for (int index = Range.Start.Row; index <= Range.End.Row; index++)
                {
                    if (indexSelector(index))
                        yield return new SheetRange(Range.Sheet, (Range.Start.Row + index, Range.Start.Col), (Range.Start.Row + index, Range.End.Col));
                }
            }
            return new SheetRangeGroup(select());
        }

    }
}
