/***************************************************************************
    File                 : opj2dat.cpp
    --------------------------------------------------------------------
    Copyright            : (C) 2008 Stefan Gerlach
                           (C) 2017 Miquel Garriga
                           (C) 2024 Dirk Lellinger (translation to C#)
    Email (use @ for *)  : stefan.gerlach*uni-konstanz.de
    Description          : Origin project converter

 ***************************************************************************/

/***************************************************************************
 *                                                                         *
 *  This program is free software; you can redistribute it and/or modify   *
 *  it under the terms of the GNU General Public License as published by   *
 *  the Free Software Foundation; either version 3 of the License, or      *
 *  (at your option) any later version.                                    *
 *                                                                         *
 *  This program is distributed in the hope that it will be useful,        *
 *  but WITHOUT ANY WARRANTY; without even the implied warranty of         *
 *  MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the          *
 *  GNU General Public License for more details.                           *
 *                                                                         *
 *   You should have received a copy of the GNU General Public License     *
 *   along with this program; if not, write to the Free Software           *
 *   Foundation, Inc., 51 Franklin Street, Fifth Floor,                    *
 *   Boston, MA  02110-1301  USA                                           *
 *                                                                         *
 ***************************************************************************/

using System;
using System.IO;

namespace Altaxo.Serialization.Origin
{
  internal class Program
  {
    private static void Main(string[] args)
    {
      if (args.Length < 1)
      {
        Console.WriteLine("Usage : opj2dat [--check-only] <file.opj>");
        Environment.Exit(-1);
      }

      Console.WriteLine($"opj2dat, Copyright (C) 2008 Stefan Gerlach, 2017 Miquel Garriga, 2024 Dirk Lellinger");

      if (args[0] == "-v")
      {
        Environment.Exit(0);
      }

      bool write_spreads = true;
      if (args.Length > 1 && args[0] == "--check-only")
      {
        write_spreads = false;
      }

      string inputfile = args[args.Length - 1];
      using var fs = new FileStream(inputfile, FileMode.Open, FileAccess.Read, FileShare.Read);
      var opj = new OriginAnyParser(fs);
      var status = opj.ParseError;
      Console.WriteLine($"Parsing status = {status}");
      if (status != 0)
      {
        Environment.Exit(-1);
      }

      Console.WriteLine($"OPJ PROJECT \"{inputfile}\" VERSION = {opj.Version}");

      Console.WriteLine($"number of datasets     = {opj.Datasets.Count}");
      Console.WriteLine($"number of spreadsheets = {opj.SpreadSheets.Count}");
      Console.WriteLine($"number of matrixes     = {opj.Matrixes.Count}");
      Console.WriteLine($"number of excels       = {opj.Excels.Count}");
      Console.WriteLine($"number of functions    = {opj.Functions.Count}");
      Console.WriteLine($"number of graphs       = {opj.Graphs.Count}");
      Console.WriteLine($"number of notes        = {opj.Notes.Count}");

      for (int s = 0; s < opj.SpreadSheets.Count; s++)
      {
        var spread = opj.SpreadSheets[s];
        var columnCount = spread.Columns.Count;
        Console.WriteLine($"Spreadsheet {s + 1}");
        Console.WriteLine($"  Name: {spread.Name}");
        Console.WriteLine($"  Label: {spread.Label}");
        Console.WriteLine($"    Columns: {columnCount}");
        for (int j = 0; j < columnCount; j++)
        {
          var column = spread.Columns[j];
          Console.WriteLine($"    Column {j + 1} : {column.Name} / type : {column.ColumnType}, rows : {spread.MaxRows}");
        }

        if (write_spreads)
        {
          string sfilename = $"{inputfile}.{s + 1}.dat";
          Console.WriteLine($"saved to {sfilename}");

          using (StreamWriter outf = new StreamWriter(sfilename))
          {
            // header
            for (int j = 0; j < columnCount; j++)
            {
              outf.Write($"{spread.Columns[j].Name}; ");
              Console.Write(spread.Columns[j].Name);
            }
            outf.WriteLine();
            Console.WriteLine("\n Data: ");
            // data
            for (int i = 0; i < (int)spread.MaxRows; i++)
            {
              for (int j = 0; j < columnCount; j++)
              {
                if (i < (int)spread.Columns[j].Data.Count)
                {
                  var value = spread.Columns[j].Data[i];
                  double v = 0.0;
                  if (value.ValueType() == Variant.VType.V_DOUBLE)
                  {
                    v = value.AsDouble();
                    if (!double.IsNaN(v))
                    {
                      outf.Write($"{v}; ");
                    }
                    else
                    {
                      outf.Write("NaN; ");
                    }
                  }
                  if (value.ValueType() == Origin.Variant.VType.V_STRING)
                  {
                    outf.Write($"{value.AsString()}; ");
                  }
                }
                else
                {
                  outf.Write("; ");
                }
              }
              outf.WriteLine();
            }
          }
        }
      }
    }
  }
}

