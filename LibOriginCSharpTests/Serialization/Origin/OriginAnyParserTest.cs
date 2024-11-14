#region Copyright

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
    public void TestOneWorksheetOneMatrix()
    {
      var fileName = Path.Combine(TestFilePath, "OneWorksheetOneMatrix.opj");
      using var str = new FileStream(fileName, FileMode.Open, FileAccess.Read, FileShare.Read);
      var reader = new OriginAnyParser(str);

      Assert.Single(reader.SpreadSheets);
      Assert.Single(reader.Matrixes);
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
      Assert.Single(matrix.Sheets);
      var sheet = matrix.Sheets[0];
      Assert.Equal(137, sheet.ColumnCount);
      Assert.Equal(179, sheet.RowCount);
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
      Assert.Single(matrix.Sheets);
      var sheet = matrix.Sheets[0];
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
      }
    }

    [Fact]
    public void TestWksDifferentNumericColumns()
    {
      var fileName = Path.Combine(TestFilePath, "WksDifferentNumericColumns.opj");
      using var str = new FileStream(fileName, FileMode.Open, FileAccess.Read, FileShare.Read);
      var reader = new OriginAnyParser(str);

      Assert.Single(reader.SpreadSheets);
      var spread = reader.SpreadSheets[0];
      Assert.Equal(10, spread.Columns.Count);

      for (int i = 0; i < spread.Columns.Count; i++)
      {
        var c = spread.Columns[0];
        Assert.Equal(0, c.BeginRow);
        Assert.Equal(2, c.EndRow);
      }
      {
        // Column A, mixed text and numeric
        var c = spread.Columns[0];
        Assert.Equal("A", c.Name);
        Assert.True(c.Data[0].IsString);
        Assert.Equal("Text", c.Data[0].AsString());
        Assert.True(c.Data[1].IsDouble);
        Assert.Equal(1.5, c.Data[1].AsDouble());
      }
      {
        // Column B (double)
        var c = spread.Columns[1];
        Assert.Equal("B", c.Name);
        Assert.True(c.Data[0].IsDouble);
        Assert.Equal(1.5, c.Data[0].AsDouble());
        Assert.True(c.Data[1].IsDouble);
        Assert.Equal(10000, c.Data[1].AsDouble());
      }
      {
        // Column C (float)
        var c = spread.Columns[2];
        Assert.Equal("C", c.Name);
        Assert.True(c.Data[0].IsDouble);
        Assert.Equal(1.5, c.Data[0].AsDouble());
        Assert.True(c.Data[1].IsDouble);
        Assert.Equal(10000, c.Data[1].AsDouble());
      }
      {
        // Column D (short)
        var c = spread.Columns[3];
        Assert.Equal("D", c.Name);
        Assert.True(c.Data[0].IsDouble);
        Assert.Equal(1, c.Data[0].AsDouble());
        Assert.True(c.Data[1].IsDouble);
        Assert.Equal(-1, c.Data[1].AsDouble());
      }
      {
        // Column E (long)
        var c = spread.Columns[4];
        Assert.Equal("E", c.Name);
        Assert.True(c.Data[0].IsDouble);
        Assert.Equal(1, c.Data[0].AsDouble());
        Assert.True(c.Data[1].IsDouble);
        Assert.Equal(-1, c.Data[1].AsDouble());
      }
      {
        // Column F (char)
        var c = spread.Columns[5];
        Assert.Equal("F", c.Name);
        Assert.True(c.Data[0].IsDouble);
        Assert.Equal(1, c.Data[0].AsDouble());
        Assert.True(c.Data[1].IsDouble);
        Assert.Equal(-1, c.Data[1].AsDouble());
      }
      {
        // Column G (byte)
        var c = spread.Columns[6];
        Assert.Equal("G", c.Name);
        Assert.True(c.Data[0].IsDouble);
        Assert.Equal(1, c.Data[0].AsDouble());
        Assert.True(c.Data[1].IsDouble);
        Assert.Equal(255, c.Data[1].AsDouble());
      }
      {
        // Column H (ushort)
        var c = spread.Columns[7];
        Assert.Equal("H", c.Name);
        Assert.True(c.Data[0].IsDouble);
        Assert.Equal(1, c.Data[0].AsDouble());
        Assert.True(c.Data[1].IsDouble);
        Assert.Equal(65535, c.Data[1].AsDouble());
      }
      {
        // Column I (ulong)
        var c = spread.Columns[8];
        Assert.Equal("I", c.Name);
        Assert.True(c.Data[0].IsDouble);
        Assert.Equal(1, c.Data[0].AsDouble());
        Assert.True(c.Data[1].IsDouble);
        Assert.Equal(uint.MaxValue, c.Data[1].AsDouble());
      }
      {
        // Column J (complex)
        var c = spread.Columns[9];
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
