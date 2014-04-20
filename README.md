# Excel Streaming Reader

If you've used [Apache POI](http://poi.apache.org) in the past to read in Excel files, you probably noticed that it's not very memory efficient. Reading in an entire workbook will cause a severe memory usage spike, which can wreak havoc on a server. 

There are plenty of good reasons for why Apache has to read in the whole workbook, but most of them have to do with the fact that the library allows you to read and write with random addresses. If (and only if) you just want to read the contents of an Excel file in a fast and memory effecient way, you probably don't need this ability. Unfortunately, there is nothing available in the POI library for reading a streaming workbook (though there is a [streaming writer](http://poi.apache.org/spreadsheet/how-to.html#sxssf)).

Well, that's what this project is for!

# Usage

This library is very specific in how it is meant to be used. You should initialize it like so:

```java
import com.thundermoose.xlsx.StreamingReader;

StreamingReader reader = StreamingReader.createReader(new File("path/to/workbook.xlsx"), 0, 100);
```

The parameters for this initialization are as follows:

1. Excel file as a java.io.File or a java.io.InputStream
2. Index of sheet to use, starting from 0
3. number of rows to cache in memory as stream is read

Once you've done this, you can then iterate through the rows and cells like so:

```java

StreamingReader reader = StreamingReader.createReader(new File("path/to/workbook.xlsx"), 0, 100);

for (Row r : reader) {
  for (Cell c : r) {
    System.out.println(c.getStringCellValue());
  }
}

```

You may access cells randomly within a row, as the entire row is cached. **However**, there is no way to randomly access rows. As this is a streaming implementation, only a small number of rows are kept in memory at any given time.

# Notes

As of right now, there is not a way of reading *directly* from an InputStream. This is entirely to do with POI's [OPCPackage.open()](http://poi.apache.org/apidocs/org/apache/poi/openxml4j/opc/OPCPackage.html) implementation. This is required to initialize the low-level stream, and unfortunately it will only perform a true stream if the input source is a `java.io.File` object. While the class *has* an overloaded method that will accept an `java.io.InputStream`, it will read the entire stream into memory.

This library provides a method to read from a stream, but it works by reading out the stream into a temporary file. Once the stream has been read out completely, it will attempt to delete the file. The behavior of this is not guaranteed, and you may end up with a lot of temp files that you don't need. If this becomes a problem, you should perform the read yourself so that you have more control over when the file is removed.

A better solution is being investigated currently.
