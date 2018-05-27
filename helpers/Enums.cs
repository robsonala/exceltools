﻿using System;

namespace exceltools.helpers
{   
    /*
    Local   Desc        Original
    ----------------------------
    0       General     0
    1       0           1
    2       0.00        2
    3       #,##0       3
    4       #,##0.00    4
    5       0%          9
    6       0.00%       10
    7       dd/mm/yyyy  N/A

    WIDTH
    -1    not set
    0       auto
    */

    public class converterSettings
    {
		public int Index { get; set; }
        public float Width { get; set; }
        public int Type { get; set; }
    }  
}
