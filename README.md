[![Build Status](https://travis-ci.com/pjfanning/excel-streaming-reader.svg?branch=master)](https://travis-ci.com/pjfanning/excel-streaming-reader)
[![Maven Central](https://maven-badges.herokuapp.com/maven-central/com.github.pjfanning/excel-streaming-reader/badge.svg)](https://maven-badges.herokuapp.com/maven-central/com.github.pjfanning/excel-streaming-reader)

# Excel Streaming Reader

This is a fork of [monitorjbl/excel-streaming-reader](https://github.com/monitorjbl/excel-streaming-reader).

This implementation supports [Apache POI](http://poi.apache.org) 5.x and only supports Java 8 and above. v2.3.x supports POI 4.x.

* [Sample](https://github.com/pjfanning/excel-streaming-reader-sample)

This implementation has some extra features
* OOXML Strict format support (see below)
* Check [Builder](https://github.com/pjfanning/excel-streaming-reader/blob/master/src/main/java/com/github/pjfanning/xlsx/StreamingReader.java#L38) implementation to see what options are available.

# Include

To use it, add this to your POM:

```
<dependencies>
  <dependency>
    <groupId>com.github.pjfanning</groupId>
    <artifactId>excel-streaming-reader</artifactId>
    <version>3.0.4</version>
  </dependency>
</dependencies>  
```

# Usage

The package name is different from the *monitorjbl/excel-streaming-reader* jar. The code is very similar.

```java
import com.github.pjfanning.xlsx.StreamingReader;

InputStream is = new FileInputStream(new File("/path/to/workbook.xlsx"));
Workbook workbook = StreamingReader.builder()
        .rowCacheSize(100)    // number of rows to keep in memory (defaults to 10)
        .bufferSize(4096)     // buffer size (in bytes) to use when reading InputStream to file (defaults to 1024)
        .open(is);            // InputStream or File for XLSX file (required)
```

Once you've done this, you can then iterate through the rows and cells like so:

```java
for (Sheet sheet : workbook){
  System.out.println(sheet.getSheetName());
  for (Row r : sheet) {
    for (Cell c : r) {
      System.out.println(c.getStringCellValue());
    }
  }
}
```

Or open a sheet by name or index:

```java
Sheet sheet = workbook.getSheet("My Sheet")
```

The StreamingWorkbook is an autocloseable resource, and it's important that you close it to free the filesystem resource it consumed. With Java 8, you can do this:

```java
try {
  InputStream is = new FileInputStream(new File("/path/to/workbook.xlsx"));
  Workbook workbook = StreamingReader.builder()
          .rowCacheSize(100)
          .bufferSize(4096)
          .open(is) {
  for (Sheet sheet : workbook){
    System.out.println(sheet.getSheetName());
    for (Row r : sheet) {
      for (Cell c : r) {
        System.out.println(c.getStringCellValue());
      }
    }
  }
}
```

You may access cells randomly within a row, as the entire row is cached. **However**, there is no way to randomly access rows. As this is a streaming implementation, only a small number of rows are kept in memory at any given time.

## Temp File Shared Strings

By default, the `/xl/sharedStrings.xml` data for your xlsx is stored in memory and this might cause memory problems.

You can use the `setUseSstTempFile(true)` option to have this data stored in a temp file (a [H2 MVStore](http://www.h2database.com/html/mvstore.html)). There is also a `setEncryptSstTempFile(true)` option if you are concerned about having the raw data in a cleartext temp file.

```java
  Workbook workbook = StreamingReader.builder()
          .rowCacheSize(100)
          .bufferSize(4096)
          .setUseSstTempFile(true)
          .setEncryptSstTempFile(true)
          .open(is)
```

# Supported Methods

Not all POI Cell and Row functions are supported. The most basic ones are (`Cell.getStringCellValue()`, `Cell.getColumnIndex()`, etc.), but don't be surprised if you get a `NotSupportedException` on the more advanced ones.

I'll try to add more support as time goes on, but some items simply can't be read in a streaming fashion. Methods that require dependent values will not have said dependencies available at the point in the stream in which they are read.

This is a brief and very generalized list of things that are not supported for reads:

* Functions
* Macros
* Styled cells (the styles are kept at the end of the ZIP file)

# OOXML Strict format

This library focuses on spreadsheets in OOXML Transitional format - despite the name, this format is more widely used. The wikipedia entry on OOXML formats has a good [description](https://en.wikipedia.org/wiki/Office_Open_XML).

* StreamingReader.Builder has `convertFromOoXmlStrict` which attempts to convert input streams from OOXML Strict format (if set to `true`) to the better supported Transitional format. This is all done in memory so this might not be a good option if you have memory limitations. This approach is now deprecated and will be removed in version 4.0.0. 
* From version 3.0.2, the standard streaming code will also try to read OOXML Strict format.
  * support is still evolving, it is recommended you use the latest available excel-streaming-reader version if you are interested in supporting OOXML Strict format 
  * `setReadComments(true)` option is not properly supported for OOXML Strict format files (https://github.com/pjfanning/excel-streaming-reader/issues/37)

# Logging

This library uses SLF4j logging. This is a rare use case, but you can plug in your logging provider and get some potentially useful output. 

# Implementation Details

This library will take a provided `InputStream` and output it to the file system. The stream is piped safely through a configurable-sized buffer to prevent large usage of memory. Once the file is created, it is then streamed into memory from the file system.

The reason for needing the stream being outputted in this manner has to do with how ZIP files work. Because the XLSX file format is basically a ZIP file, it's not possible to find all of the entries without reading the entire InputStream.

This is a problem that can't really be gotten around for POI, as it needs a complete list of ZIP entries. The default implementation of reading from an `InputStream` in POI is to read the entire stream directly into memory. This library works by reading out the stream into a temporary file. As part of the auto-close action, the temporary file is deleted.

If you need more control over how the file is created/disposed of, there is an option to initialize the library with a `java.io.File`. This file will not be written to or removed:

```java
File f = new File("/path/to/workbook.xlsx");
Workbook workbook = StreamingReader.builder()
        .rowCacheSize(100)    
        .bufferSize(4096)     
        .open(f);
```

This library will ONLY work with XLSX files. The older XLS format is not capable of being streamed.
