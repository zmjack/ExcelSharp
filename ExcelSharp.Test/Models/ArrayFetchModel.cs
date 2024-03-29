﻿using ExcelSharp.NPOI;

namespace ExcelSharp.Test.Models;

public class ArrayFetchModel
{
    public string Name { get; set; }

    [Fetch(10)]
    public int?[] Numbers { get; set; }

    public int? Total { get; set; }
}
