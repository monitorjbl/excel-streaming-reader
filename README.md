![Build Status](https://github.com/pjfanning/excel-streaming-reader/actions/workflows/ci.yml/badge.svg)
[![Maven Central](https://maven-badges.herokuapp.com/maven-central/com.github.pjfanning/excel-streaming-reader/badge.svg)](https://maven-badges.herokuapp.com/maven-central/com.github.pjfanning/excel-streaming-reader)

# Excel Streaming Reader

This is a fork of [monitorjbl/excel-streaming-reader](https://github.com/monitorjbl/excel-streaming-reader).

This implementation supports [Apache POI](http://poi.apache.org) 5.x and only supports Java 8 and above. v2.3.x supports POI 4.x.

* [Sample](https://github.com/pjfanning/excel-streaming-reader-sample)

This implementation has some extra features
* OOXML Strict format support (see below)
* More methods are implemented. Some require that features are enabled in the StreamingReader.Builder instance because they might have an additional overhead.
* Check [Builder](https://pjfanning.github.io/excel-streaming-reader/javadocs/3.3.0/com/github/pjfanning/xlsx/StreamingReader.Builder.html) implementation to see what options are available.

## Used By
* [Apache Drill](https://drill.apache.org/)
* [Apache Linkis](https://linkis.apache.org/)
* [Spark-Excel](https://github.com/crealytics/spark-excel)
* [Sirius Web](https://sirius-lib.net/#web-features)

# Include

To use it, add this to your POM:

```
<dependencies>
  <dependency>
    <groupId>com.github.pjfanning</groupId>
    <artifactId>excel-streaming-reader</artifactId>
    <version>3.4.1</version>
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
try (
        InputStream is = new FileInputStream(new File("/path/to/workbook.xlsx"));
        Workbook workbook = StreamingReader.builder()
          .rowCacheSize(100)
          .bufferSize(4096)
          .open(is)
){
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

## Javadocs

* [latest](https://pjfanning.github.io/excel-streaming-reader/)
* [v3.1.6](https://pjfanning.github.io/excel-streaming-reader/javadocs/3.1.6/)

## Temp File Shared Strings

By default, the `/xl/sharedStrings.xml` data for your xlsx is stored in memory and this might cause memory problems.

You can use the `setUseSstTempFile(true)` option to have this data stored in a temp file (a [H2 MVStore](http://www.h2database.com/html/mvstore.html)). There is also a `setEncryptSstTempFile(true)` option if you are concerned about having the raw data in a cleartext temp file.

```java
  Workbook workbook = StreamingReader.builder()
          .setUseSstTempFile(true)
          .setEncryptSstTempFile(false)
          .fullFormatRichText(true) //if you want the rich text formatting as well as the text
          .open(is);
```

## Temp File Comments

As with shared strings, comments are stored in a separate part of the xlsx file and by default,
excel-streaming-reader does not read them. You can configure excel-streaming-reader to read them and
choose whether you want them stored in memory or in a temp file while reading the xlsx file.

```java
  Workbook workbook = StreamingReader.builder()
          .setReadComments(true)
          .setUseCommentsTempFile(true)
          .setEncryptCommentsTempFile(false)
          .fullFormatRichText(true) //if you want the rich text formatting as well as the text
          .open(is);
```

## Reading Very Large Excel Files

excel-streaming-reader uses some Apache POI code under the hood. That code uses memory and/or
temp files to store temporary data while it processes the xlsx. With very large files, you will probably
want to favour using temp files.

With `StreamingReader.builder()`, do not set `setAvoidTempFiles(true)`. You should also consider, tuning
[POI settings](https://poi.apache.org/components/configuration.html) too. In particular,
consider setting these properties:

```java
  org.apache.poi.openxml4j.util.ZipInputStreamZipEntrySource.setThresholdBytesForTempFiles(16384); //16KB
  org.apache.poi.openxml4j.opc.ZipPackage.setUseTempFilePackageParts(true);
```

# Supported Methods

Not all POI Cell and Row functions are supported. The most basic ones are (`Cell.getStringCellValue()`, `Cell.getColumnIndex()`, etc.), but don't be surprised if you get a `NotSupportedException` on the more advanced ones.

I'll try to add more support as time goes on, but some items simply can't be read in a streaming fashion. Methods that require dependent values will not have said dependencies available at the point in the stream in which they are read.

This is a brief and very generalized list of things that are not supported for reads:

* Recalculating Formulas - you will get values that Excel cached in the xlsx when the file was saved
* Macros

# OOXML Strict format

This library focuses on spreadsheets in OOXML Transitional format - despite the name, this format is more widely used. The wikipedia entry on OOXML formats has a good [description](https://en.wikipedia.org/wiki/Office_Open_XML).

* From version 3.0.2, the standard streaming code will also try to read OOXML Strict format.
  * support is still evolving, it is recommended you use the latest available excel-streaming-reader version if you are interested in supporting OOXML Strict format
* Version 3.2.0 drops StreamingReader.Builder `convertFromOoXmlStrict` option (previously deprecated) as this is supported by default now.

# Logging

This library uses [SLF4j](http://www.slf4j.org/) logging. This is a rare use case, but you can plug in your logging provider and get some potentially useful output. POI 5.1.0 switched to [Log4j 2.x](https://logging.apache.org/log4j/2.x/) for logging. If you need logs from both libraries, you will need to use one of the bridge jars to map slf4j to log4j or vice versa.

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
