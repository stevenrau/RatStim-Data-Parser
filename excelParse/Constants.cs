/**************************************************************************\
Module Name:   Constants.cs 
Project:       excelParse
Author:        Steven Rau

This file is is not really a class as much as it is a container for 
any constant values needed thoughout the project.
\***************************************************************************/

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace RatStim
{
    public static class Constants
    {
        //We will be parsing .csv files with 17 columns
        const int NUM_CSV_COLS = 17;

        //The different stim values for each rat
        public static readonly string[] stims= {"No_Stim", "p120", 
                                         "PP12(140ms)P120", "PP12(30ms)P120", "PP12(50ms)P120", "PP12(80ms)P120", "PP12alone",
                                         "PP3(140ms)P120", "PP3(30ms)P120", "PP3(50ms)P120", "PP3(80ms)P120", "PP3alone",
                                         "PP6(140ms)P120", "PP6(30ms)P120", "PP6(50ms)P120", "PP6(80ms)P120", "PP6alone"};
    }
}
