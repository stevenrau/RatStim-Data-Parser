/**************************************************************************\
Module Name:   RatById.cs 
Project:       excelParse
Author:        Steven Rau

This class acts as a container to store rat entries with shared IDs in
a list. Each uique rat ID will have its own RatById class with all of
its corresponding entries.
\***************************************************************************/

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace RatStim
{
    class RatById
    {
        public List<Entry> entries;                //A list of entries for each unique rat ID
        public string id;                          //The unique rat ID
        public Dictionary<string, double> avgs;   //The avgs of each stim value for this rat. The key is the stim value

        public RatById(string newId)
        {
            id = newId;
            entries = new List<Entry>();
            avgs = new Dictionary<string, double>();
        }

        /**
         * Adds the avg entry with the key stimVal to the avgs dictionary. Average should 
         * never be negative sinec these are avg response times.
         *
         * @param  stimVal  The string representing the stimuus value avg we are storing
         * @param  avg      The calcuated avg value for the given stimulus value
         * 
         * @return  FAILURE  For invalid parameter, such as null string or negative avg
         *          SUCCESS  For successful add
         */
        public int addAvg(string stimVal, double avg)
        {
            if (null == stimVal || 0 > avg)
            {
                return Constants.FAILURE;
            }

            avgs.Add(stimVal, avg);

            return Constants.SUCCESS;
        }

        /**
         * Gets the avg value for the given stimVal from the avgs dictionary, if it exists
         * 
         * @param  stimVal  The string key for the stimulus avg we want to get
         * 
         * @return  FAILURE  When trying to get an avg with a key that DNE in the dictionary
         *          avg      The avg value from the dictionary that corresponds to the stimVal key on success
         */
        public double getAvg(string stimVal)
        {
            double avg;
            if (avgs.TryGetValue(stimVal, out avg))
            {
                // Key was in dictionary; "avg" contains corresponding value
                return avg;
            }
            else
            {
                // Key wasn't in dictionary; "avg" is now 0
                return (double)Constants.FAILURE;
            }
        }
    }
}
