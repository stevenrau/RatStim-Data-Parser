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

namespace excelParse
{
    class RatById
    {

        public List<Entry> entries;
        public int id;

        public RatById(int newId)
        {
            id = newId;
            entries = new List<Entry>();
        }
    }
}
