﻿#region Copyright

/////////////////////////////////////////////////////////////////////////////
//    Altaxo:  a data processing and data plotting program
//    Copyright (C) 2002-2024 Dr. Dirk Lellinger
//
//    This program is free software; you can redistribute it and/or modify
//    it under the terms of the GNU General Public License as published by
//    the Free Software Foundation; either version 2 of the License, or
//    (at your option) any later version.
//
//    This program is distributed in the hope that it will be useful,
//    but WITHOUT ANY WARRANTY; without even the implied warranty of
//    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
//    GNU General Public License for more details.
//
//    You should have received a copy of the GNU General Public License
//    along with this program; if not, write to the Free Software
//    Foundation, Inc., 675 Mass Ave, Cambridge, MA 02139, USA.
//
/////////////////////////////////////////////////////////////////////////////

#endregion Copyright

using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Xunit;

namespace Altaxo.Serialization.Origin.Tests
{
  public class OriginFile_Test
  {
    public string TestFilePath => Path.Combine(AppDomain.CurrentDomain.BaseDirectory, @"Serialization\Origin\TestFiles");


    /// <summary>
    /// Tests if all .opj files in the test folder are readable.
    /// </summary>
    [Fact]
    public void Test_AllFilesReadable()
    {
      var opjFiles = new DirectoryInfo(TestFilePath).GetFiles("*.opj");
      Assert.NotEmpty(opjFiles);
      foreach (var file in opjFiles)
      {
        using var str = new FileStream(file.FullName, FileMode.Open, FileAccess.Read, FileShare.Read);
        var reader = new OriginAnyParser(str);
      }
    }

    /// <summary>
    /// Tests if .opj files in additional folders (not included in the test folder) are readable.
    /// This test expect that there is a file 'AdditionalFoldersWithOpjFiles.txt' in the test file folder,
    /// which you have to create (this file is not tracked in Git).
    /// This text file should contain additional folder names (each folder name located on a separate line) which contains .opj files.
    /// The folders are searched then recursively for .opj files to test.
    /// The test will not fail if there is no such 'AdditionalFoldersWithOpjFiles.txt' file,
    /// or if the additional folders do not contain .opj files.
    /// </summary>
    [Fact]
    public void Test_AdditionalFilesReadable()
    {
      var listOfFailedFiles = new List<(FileInfo fileInfo, Exception exception)>();
      var listOfCorruptedFiles = new List<(FileInfo fileInfo, Exception exception)>();

      void TestFolder(DirectoryInfo folder)
      {
        if (folder.Name == "$RECYCLE.BIN")
          return;

        FileInfo[] opjFiles;
        try
        {
          opjFiles = folder.GetFiles("*.opj", SearchOption.TopDirectoryOnly);
        }
        catch (Exception ex)
        {
          return;
        }
        foreach (var file in opjFiles)
        {
          TestFile(file);
        }
        var subFolders = folder.GetDirectories();
        foreach (var subFolder in subFolders)
        {
          TestFolder(subFolder);
        }
      }
      void TestFile(FileInfo file)
      {
        using var str = new FileStream(file.FullName, FileMode.Open, FileAccess.Read, FileShare.Read);
        (int FileVersion, int NewFileVersion, int BuildVersion, bool IsOpjuFile, string? Error) version;
        try
        {
          version = OriginAnyParser.ReadFileVersion(str);
        }
        catch (Exception ex)
        {
          listOfCorruptedFiles.Add((file, ex));
          return;
        }
        if (version.FileVersion >= 400)
        {
          try
          {
            var reader = new OriginAnyParser(str);
          }
          catch (EndOfStreamException ex)
          {
            listOfCorruptedFiles.Add((file, ex));
          }
          catch (Exception ex)
          {
            listOfFailedFiles.Add((file, ex));
          }
        }
      }

      var additionalFoldersFile = new FileInfo(Path.Combine(TestFilePath, "AdditionalFoldersWithOpjFiles.txt"));

      if (additionalFoldersFile.Exists) // Note that it is not an error if there is no such file
      {
        using var fr = new StreamReader(additionalFoldersFile.FullName, true);
        string? line;
        while (null != (line = fr.ReadLine()))
        {
          line = line.Trim();
          if (!string.IsNullOrEmpty(line))
          {
            var file = new FileInfo(line.Trim().Trim('"'));
            if (file.Exists)
            {
              TestFile(file);
              continue;
            }
            var folder = new DirectoryInfo(line);
            if (folder.Exists)
            {
              TestFolder(folder);
              continue;
            }
          }
        }
      }
      if (listOfFailedFiles.Count > 0 || listOfCorruptedFiles.Count > 0)
      {
        // Set a break point here to inspect the list of failed files
      }

      if (listOfCorruptedFiles.Count > 0 || listOfFailedFiles.Count > 0)
      {
        var stb = new StringBuilder();
        if (listOfCorruptedFiles.Count > 0)
        {
          stb.AppendLine("List of corrupted files:");
          foreach (var file in listOfCorruptedFiles)
          {
            stb.AppendLine($"\"{file.fileInfo.FullName}\"\t{file.exception.Message}");
          }
        }

        if (listOfFailedFiles.Count > 0)
        {

          stb.AppendLine("List of failed files:");
          foreach (var file in listOfFailedFiles)
          {
            stb.AppendLine($"\"{file.fileInfo.FullName}\"\t{file.exception.Message}");
          }
        }
        Assert.Fail(stb.ToString());
      }

    }

    [Fact]
    public void TestMatrix137x179()
    {
      var fileName = Path.Combine(TestFilePath, "Matrix137x179.opj");
      using var str = new FileStream(fileName, FileMode.Open, FileAccess.Read, FileShare.Read);
      var reader = new OriginAnyParser(str);

      Assert.Empty(reader.SpreadSheets);
      Assert.Single(reader.Matrixes);
      var matrix = reader.Matrixes[0];
      Assert.Equal("MBook1", matrix.Name);
      Assert.Single(matrix.Sheets);
      var sheet = matrix.Sheets[0];
      Assert.Equal("MSheet1", sheet.Name);
      Assert.Equal(137, sheet.ColumnCount);
      Assert.Equal(179, sheet.RowCount);
      Assert.Equal(13, sheet.X1);
      Assert.Equal(21, sheet.X2);
      Assert.Equal(17, sheet.Y1);
      Assert.Equal(37, sheet.Y2);
      Assert.Equal(0, sheet[0, 0]);
      Assert.Equal(1, sheet[0, 1]);
      Assert.Equal(2, sheet[0, 2]);
      Assert.Equal(3, sheet[0, 3]);
      Assert.Equal(10, sheet[1, 0]);
      Assert.Equal(20, sheet[2, 0]);
      Assert.Equal(30, sheet[3, 0]);
    }

    [Fact]
    public void TestMatrix2x3OfDifferentTypes()
    {
      var fileName = Path.Combine(TestFilePath, "Matrix2x3OfDifferentTypes.opj");
      using var str = new FileStream(fileName, FileMode.Open, FileAccess.Read, FileShare.Read);
      var reader = new OriginAnyParser(str);

      Assert.Empty(reader.SpreadSheets);
      Assert.Equal(9, reader.Matrixes.Count);

      foreach (var matrix in reader.Matrixes)
      {
        var sheet = matrix.Sheets[0];
        Assert.Equal("MSheet1", sheet.Name);
        Assert.Equal(11, sheet[0, 0]);
        Assert.Equal(12, sheet[0, 1]);
        Assert.Equal(21, sheet[1, 0]);
        Assert.Equal(22, sheet[1, 1]);
        Assert.Equal(31, sheet[2, 0]);
        Assert.Equal(32, sheet[2, 1]);

        if (matrix.Name == "MComplex")
        {
          Assert.Equal(0.11, sheet.ImaginaryPart(0, 0));
          Assert.Equal(0.12, sheet.ImaginaryPart(0, 1));
          Assert.Equal(0.21, sheet.ImaginaryPart(1, 0));
          Assert.Equal(0.22, sheet.ImaginaryPart(1, 1));
          Assert.Equal(0.31, sheet.ImaginaryPart(2, 0));
          Assert.Equal(0.32, sheet.ImaginaryPart(2, 1));
        }

        switch (matrix.Name)
        {
          case "MDouble":
          case "MReal":
          case "MShort":
          case "MLong":
          case "MChar":
          case "MByte":
          case "MUShort":
          case "MULong":
          case "MComplex":
            break;
          default:
            Assert.Fail($"Unexpected name of matrix: {matrix.Name}");
            break;
        }
      }
    }

    [Fact]
    public void TestMatrix71x29()
    {
      var fileName = Path.Combine(TestFilePath, "Matrix71x29.opj");
      using var str = new FileStream(fileName, FileMode.Open, FileAccess.Read, FileShare.Read);
      var reader = new OriginAnyParser(str);

      Assert.Empty(reader.SpreadSheets);
      Assert.Single(reader.Matrixes);
      var matrix = reader.Matrixes[0];
      Assert.Equal("MBook1", matrix.Name);
      Assert.Single(matrix.Sheets);
      var sheet = matrix.Sheets[0];
      Assert.Equal("MSheet1", sheet.Name);
      Assert.Equal(71, sheet.ColumnCount);
      Assert.Equal(29, sheet.RowCount);
      Assert.Equal(1, sheet[0, 0]);
      Assert.Equal(2, sheet[0, 1]);
      Assert.Equal(3, sheet[0, 2]);
      Assert.Equal(10, sheet[1, 0]);
      Assert.Equal(20, sheet[2, 0]);
      Assert.Equal(30, sheet[3, 0]);

      // test the coordinates
      Assert.Equal(5329, sheet.X1);
      Assert.Equal(9999, sheet.X2);
      Assert.Equal(731, sheet.Y1);
      Assert.Equal(999, sheet.Y2);
    }

    [Fact]
    public void TestOneWorksheet()
    {
      var fileName = Path.Combine(TestFilePath, "OneWorksheet.opj");
      using var str = new FileStream(fileName, FileMode.Open, FileAccess.Read, FileShare.Read);
      var reader = new OriginAnyParser(str);

      Assert.Single(reader.SpreadSheets);
      Assert.Empty(reader.Matrixes);

      var spreadSheet = reader.SpreadSheets[0];
      Assert.Equal("Book1", spreadSheet.Name);
      Assert.Equal(2, spreadSheet.Columns.Count);

      var c = spreadSheet.Columns[0];
      Assert.Equal(0, c.BeginRow);
      Assert.Equal(3, c.EndRow);
      Assert.Equal(1, c.Data[0].AsDouble());
      Assert.Equal(2, c.Data[1].AsDouble());
      Assert.Equal(3, c.Data[2].AsDouble());
      Assert.Equal("Leitfähigkeit", c.LongName);
      Assert.Equal("S", c.Units);
      Assert.Empty(c.Comments);

      c = spreadSheet.Columns[1];
      Assert.Equal(0, c.BeginRow);
      Assert.Equal(3, c.EndRow);
      Assert.Equal(4, c.Data[0].AsDouble());
      Assert.Equal(5, c.Data[1].AsDouble());
      Assert.Equal(6, c.Data[2].AsDouble());
      Assert.Equal("Temperatur", c.LongName);
      Assert.Equal("°C", c.Units);
      Assert.Empty(c.Comments);
    }

    [Fact]
    public void TestOneWorksheetOneMatrix()
    {
      var fileName = Path.Combine(TestFilePath, "OneWorksheetOneMatrix.opj");
      using var str = new FileStream(fileName, FileMode.Open, FileAccess.Read, FileShare.Read);
      var reader = new OriginAnyParser(str);

      Assert.Single(reader.SpreadSheets);
      Assert.Single(reader.Matrixes);
      var spreadSheet = reader.SpreadSheets[0];
      Assert.Equal("Book1", spreadSheet.Name);
      Assert.Equal(3, spreadSheet.Columns.Count);
      for (int c = 0; c < 3; ++c)
      {
        var column = spreadSheet.Columns[c];
        Assert.Equal(0, column.BeginRow);
        Assert.Equal(5, column.EndRow);
        for (int r = 0; r < 5; ++r)
        {
          Assert.Equal(c + 1 + 3 * r, column.Data[r].AsDouble());
        }
      }

      var matrix = reader.Matrixes[0];
      Assert.Equal("MBook1", matrix.Name);
      Assert.Single(matrix.Sheets);
      var matrixSheet = matrix.Sheets[0];
      Assert.Equal("MSheet1", matrixSheet.Name);
      Assert.Equal(3, matrixSheet.ColumnCount);
      Assert.Equal(5, matrixSheet.RowCount);

      for (int c = 0; c < 3; ++c)
      {
        for (int r = 0; r < 5; ++r)
        {
          Assert.Equal(c + 1 + 3 * r, matrixSheet[r, c]);
        }
      }
      Assert.Equal(1, matrixSheet.X1);
      Assert.Equal(3, matrixSheet.X2);
      Assert.Equal(1, matrixSheet.Y1);
      Assert.Equal(5, matrixSheet.Y2);
    }

    [Fact]
    public void TestOneWorksheetWithTwoSheets()
    {
      var fileName = Path.Combine(TestFilePath, "OneWorksheetWithTwoSheets.opj");
      using var str = new FileStream(fileName, FileMode.Open, FileAccess.Read, FileShare.Read);
      var reader = new OriginAnyParser(str);

      Assert.Empty(reader.SpreadSheets);
      Assert.Empty(reader.Matrixes);
      Assert.Single(reader.Excels);

      var excel = reader.Excels[0];
      Assert.Equal("Book1", excel.Name);
      Assert.Equal(2, excel.Sheets.Count);

      var spreadSheet = excel.Sheets[0];
      Assert.Equal(2, spreadSheet.Columns.Count);
      Assert.Equal("Sheet1", spreadSheet.Name);
      var c = spreadSheet.Columns[0];
      Assert.Equal("A", c.Name);
      Assert.Equal(0, c.BeginRow);
      Assert.Equal(3, c.EndRow);
      Assert.Equal(1, c.Data[0].AsDouble());
      Assert.Equal(2, c.Data[1].AsDouble());
      Assert.Equal(3, c.Data[2].AsDouble());
      Assert.Empty(c.LongName);
      Assert.Empty(c.Units);
      Assert.Empty(c.Comments);

      c = spreadSheet.Columns[1];
      Assert.Equal("B", c.Name);
      Assert.Equal(0, c.BeginRow);
      Assert.Equal(3, c.EndRow);
      Assert.Equal(4, c.Data[0].AsDouble());
      Assert.Equal(5, c.Data[1].AsDouble());
      Assert.Equal(6, c.Data[2].AsDouble());
      Assert.Empty(c.LongName);
      Assert.Empty(c.Units);
      Assert.Empty(c.Comments);

      spreadSheet = excel.Sheets[1];
      Assert.Equal(2, spreadSheet.Columns.Count);
      Assert.Equal("Sheet2", spreadSheet.Name);

      c = spreadSheet.Columns[0];
      Assert.Equal("A", c.Name);
      Assert.True(c.BeginRow <= 4);
      Assert.Equal(7, c.EndRow);
      for (int r = c.BeginRow; r < 4; ++r)
      {
        var d = c.Data[r];
        Assert.True(d.IsDoubleOrNaN);
        Assert.True(double.IsNaN(d.AsDouble()));
      }
      Assert.Equal(7, c.Data[4 - c.BeginRow].AsDouble());
      Assert.Equal(8, c.Data[5 - c.BeginRow].AsDouble());
      Assert.Equal(9, c.Data[6 - c.BeginRow].AsDouble());
      Assert.Empty(c.LongName);
      Assert.Empty(c.Units);
      Assert.Empty(c.Comments);

      c = spreadSheet.Columns[1];
      Assert.Equal("B", c.Name);
      Assert.True(c.BeginRow <= 4);
      Assert.Equal(7, c.EndRow);
      for (int r = c.BeginRow; r < 4; ++r)
      {
        var d = c.Data[r];
        Assert.True(d.IsDoubleOrNaN);
        Assert.True(double.IsNaN(d.AsDouble()));
      }
      Assert.Equal(10, c.Data[4 - c.BeginRow].AsDouble());
      Assert.Equal(11, c.Data[5 - c.BeginRow].AsDouble());
      Assert.Equal(12, c.Data[6 - c.BeginRow].AsDouble());
      Assert.Empty(c.LongName);
      Assert.Empty(c.Units);
      Assert.Empty(c.Comments);
    }


    [Fact]
    public void TestTwoWorksheets()
    {
      var fileName = Path.Combine(TestFilePath, "TwoWorksheets.opj");
      using var str = new FileStream(fileName, FileMode.Open, FileAccess.Read, FileShare.Read);
      var reader = new OriginAnyParser(str);

      Assert.Equal(2, reader.SpreadSheets.Count);
      Assert.Empty(reader.Matrixes);
      Assert.Empty(reader.Excels);

      var spreadSheet = reader.SpreadSheets[0];
      Assert.Equal("Book1", spreadSheet.Name);
      Assert.Equal(2, spreadSheet.Columns.Count);

      var c = spreadSheet.Columns[0];
      Assert.Equal(0, c.BeginRow);
      Assert.Equal(3, c.EndRow);
      Assert.Equal(1, c.Data[0].AsDouble());
      Assert.Equal(2, c.Data[1].AsDouble());
      Assert.Equal(3, c.Data[2].AsDouble());
      Assert.Equal("Leitfähigkeit", c.LongName);
      Assert.Equal("S", c.Units);
      Assert.Empty(c.Comments);

      c = spreadSheet.Columns[1];
      Assert.Equal(0, c.BeginRow);
      Assert.Equal(3, c.EndRow);
      Assert.Equal(4, c.Data[0].AsDouble());
      Assert.Equal(5, c.Data[1].AsDouble());
      Assert.Equal(6, c.Data[2].AsDouble());
      Assert.Equal("Temperatur", c.LongName);
      Assert.Equal("°C", c.Units);
      Assert.Empty(c.Comments);

      spreadSheet = reader.SpreadSheets[1];
      Assert.Equal("Book2", spreadSheet.Name);
      Assert.Equal(2, spreadSheet.Columns.Count);

      c = spreadSheet.Columns[0];
      Assert.Equal(0, c.BeginRow);
      Assert.Equal(3, c.EndRow);
      Assert.Equal("1,1", c.Data[0].AsString());
      Assert.Equal("2,2", c.Data[1].AsString());
      Assert.Equal("3,3", c.Data[2].AsString());
      Assert.Equal("LN1", c.LongName);
      Assert.Equal("Siemens", c.Units);
      Assert.Equal("Kommentar1", c.Comments);

      c = spreadSheet.Columns[1];
      Assert.Equal(0, c.BeginRow);
      Assert.Equal(3, c.EndRow);
      Assert.Equal("4,4", c.Data[0].AsString());
      Assert.Equal(5.5, c.Data[1].AsDouble());
      Assert.Equal("6,6", c.Data[2].AsString());
      Assert.Equal("LN2", c.LongName);
      Assert.Equal("°C", c.Units);
      Assert.Equal("Kommentar2", c.Comments);


    }

    [Fact]
    public void TestWksDifferentNumericColumns()
    {
      var fileName = Path.Combine(TestFilePath, "WksDifferentNumericColumns.opj");
      using var str = new FileStream(fileName, FileMode.Open, FileAccess.Read, FileShare.Read);
      var reader = new OriginAnyParser(str);

      Assert.Single(reader.SpreadSheets);
      var spreadSheet = reader.SpreadSheets[0];
      Assert.Equal("Book1", spreadSheet.Name);
      Assert.Equal(10, spreadSheet.Columns.Count);

      for (int i = 0; i < spreadSheet.Columns.Count; i++)
      {
        var c = spreadSheet.Columns[0];
        Assert.Equal(0, c.BeginRow);
        Assert.Equal(2, c.EndRow);
      }
      {
        // Column A, mixed text and numeric
        var c = spreadSheet.Columns[0];
        Assert.Equal("A", c.Name);
        Assert.True(c.Data[0].IsString);
        Assert.Equal("Text", c.Data[0].AsString());
        Assert.True(c.Data[1].IsDouble);
        Assert.Equal(1.5, c.Data[1].AsDouble());
      }
      {
        // Column B (double)
        var c = spreadSheet.Columns[1];
        Assert.Equal("B", c.Name);
        Assert.True(c.Data[0].IsDouble);
        Assert.Equal(1.5, c.Data[0].AsDouble());
        Assert.True(c.Data[1].IsDouble);
        Assert.Equal(10000, c.Data[1].AsDouble());
      }
      {
        // Column C (float)
        var c = spreadSheet.Columns[2];
        Assert.Equal("C", c.Name);
        Assert.True(c.Data[0].IsDouble);
        Assert.Equal(1.5, c.Data[0].AsDouble());
        Assert.True(c.Data[1].IsDouble);
        Assert.Equal(10000, c.Data[1].AsDouble());
      }
      {
        // Column D (short)
        var c = spreadSheet.Columns[3];
        Assert.Equal("D", c.Name);
        Assert.True(c.Data[0].IsDouble);
        Assert.Equal(1, c.Data[0].AsDouble());
        Assert.True(c.Data[1].IsDouble);
        Assert.Equal(-1, c.Data[1].AsDouble());
      }
      {
        // Column E (long)
        var c = spreadSheet.Columns[4];
        Assert.Equal("E", c.Name);
        Assert.True(c.Data[0].IsDouble);
        Assert.Equal(1, c.Data[0].AsDouble());
        Assert.True(c.Data[1].IsDouble);
        Assert.Equal(-1, c.Data[1].AsDouble());
      }
      {
        // Column F (char)
        var c = spreadSheet.Columns[5];
        Assert.Equal("F", c.Name);
        Assert.True(c.Data[0].IsDouble);
        Assert.Equal(1, c.Data[0].AsDouble());
        Assert.True(c.Data[1].IsDouble);
        Assert.Equal(-1, c.Data[1].AsDouble());
      }
      {
        // Column G (byte)
        var c = spreadSheet.Columns[6];
        Assert.Equal("G", c.Name);
        Assert.True(c.Data[0].IsDouble);
        Assert.Equal(1, c.Data[0].AsDouble());
        Assert.True(c.Data[1].IsDouble);
        Assert.Equal(255, c.Data[1].AsDouble());
      }
      {
        // Column H (ushort)
        var c = spreadSheet.Columns[7];
        Assert.Equal("H", c.Name);
        Assert.True(c.Data[0].IsDouble);
        Assert.Equal(1, c.Data[0].AsDouble());
        Assert.True(c.Data[1].IsDouble);
        Assert.Equal(65535, c.Data[1].AsDouble());
      }
      {
        // Column I (ulong)
        var c = spreadSheet.Columns[8];
        Assert.Equal("I", c.Name);
        Assert.True(c.Data[0].IsDouble);
        Assert.Equal(1, c.Data[0].AsDouble());
        Assert.True(c.Data[1].IsDouble);
        Assert.Equal(uint.MaxValue, c.Data[1].AsDouble());
      }
      {
        // Column J (complex)
        var c = spreadSheet.Columns[9];
        Assert.Equal("J", c.Name);
        Assert.True(c.Data[0].IsDouble);
        Assert.Equal(1.5, c.Data[0].AsDouble());
        Assert.True(c.Data[1].IsDouble);
        Assert.Equal(-1.5, c.Data[1].AsDouble());
        Assert.NotNull(c.ImaginaryData);
        Assert.Equal(0.5, c.ImaginaryData[0]);
        Assert.Equal(3.5, c.ImaginaryData[1]);
      }
    }
  }
}
