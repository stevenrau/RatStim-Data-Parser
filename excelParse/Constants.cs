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

        //Return codes to be used throughout the project
        public const int SUCCESS = 1;
        public const int FAILURE = -1;

        //Columns for the corresponding values in the Master file
        public const int RAT_ID = 1;
        public const int STRAIN = 2;
        public const int WEIGHT = 3;
        public const int P120_BEFORE = 4;
        public const int P120_DURING = 5;
        public const int P120_AFTER = 6;
        public const int NO_STIM = 7;
        public const int PP3_ALONE = 8;
        public const int PP6_ALONE = 9;
        public const int PP12_ALONE = 10;
        public const int PP3_30 = 11;
        public const int PP6_30 = 12;
        public const int PP12_30 = 13;
        public const int PP3_50 = 14;
        public const int PP6_50 = 15;
        public const int PP12_50 = 16;
        public const int PP3_80 = 17;
        public const int PP6_80 = 18;
        public const int PP12_80 = 19;
        public const int PP3_140 = 20;
        public const int PP6_140 = 21;
        public const int PP12_140 = 22;

        //String representations of the stimulus values
        public const string P120_BEFORE_STR = "p120_before";
        public const string P120_DURING_STR = "p120_during";
        public const string P120_AFTER_STR = "p120_after";
        public const string NO_STIM_STR = " No_Stim";
        public const string PP3_ALONE_STR = " PP3alone";
        public const string PP6_ALONE_STR = " PP6alone";
        public const string PP12_ALONE_STR = " PP12alone";
        public const string PP3_30_STR = " PP3(30ms)P120";
        public const string PP6_30_STR = " PP6(30ms)P120";
        public const string PP12_30_STR = " PP12(30ms)P120";
        public const string PP3_50_STR = " PP3(50ms)P120";
        public const string PP6_50_STR = " PP6(50ms)P120";
        public const string PP12_50_STR = " PP12(50ms)P120";
        public const string PP3_80_STR = " PP3(80ms)P120";
        public const string PP6_80_STR = " PP6(80ms)P120";
        public const string PP12_80_STR = " PP12(80ms)P120";
        public const string PP3_140_STR = " PP3(140ms)P120";
        public const string PP6_140_STR = " PP6(140ms)P120";
        public const string PP12_140_STR = " PP12(140ms)P120";


        //The different stim values for each rat
        public static readonly string[] stims= {"No_Stim", "p120", 
                                         "PP12(140ms)P120", "PP12(30ms)P120", "PP12(50ms)P120", "PP12(80ms)P120", "PP12alone",
                                         "PP3(140ms)P120", "PP3(30ms)P120", "PP3(50ms)P120", "PP3(80ms)P120", "PP3alone",
                                         "PP6(140ms)P120", "PP6(30ms)P120", "PP6(50ms)P120", "PP6(80ms)P120", "PP6alone"};
    }
}
