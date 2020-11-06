package com.github.pjfanning.xlsx.impl.ooxml;

import org.apache.poi.util.Beta;

import javax.xml.stream.XMLStreamException;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.zip.ZipEntry;
import java.util.zip.ZipInputStream;
import java.util.zip.ZipOutputStream;

@Beta
public class OoXmlStrictConverterInputStream extends InputStream {
    private final ZipInputStream input;
    private final ByteArrayOutputStreamExposed buffer = new ByteArrayOutputStreamExposed();
    private final ZipOutputStream output;
    private ReadState readState;
    private ZipEntry zipEntry;
    private byte[] nonXmlEntryReadBuffer = new byte[4096];
    private OoXmlStrictConverter xmlConverter;
    private int readIndex = 0;

    public OoXmlStrictConverterInputStream(InputStream ooXmlStrictInput) {
        this.input = new ZipInputStream(ooXmlStrictInput);
        this.output = new ZipOutputStream(buffer);
        this.readState = ReadState.BEFORE_ZIP_ENTRY;
    }

    @Override
    public int read() throws IOException {

        while (!hasBytesToRead()) {
            if (readState == ReadState.END_STATE) {
                return -1;
            }
            keepReadingFromInput();
        }

        ++readIndex;
        return buffer.getBytes()[readIndex - 1];
    }

    @Override
    public int read(byte b[], int off, int len) throws IOException {
        if (b == null) {
            throw new NullPointerException();
        } else if (off < 0 || len < 0 || len > b.length - off) {
            throw new IndexOutOfBoundsException();
        } else if (len == 0) {
            return 0;
        }

        while (!hasBytesToRead()) {
            if (readState == ReadState.END_STATE) {
                return -1;
            }
            keepReadingFromInput();
        }

        int readBytes = 0;

        while (readBytes < len && hasBytesToRead()) {
            b[off + readBytes] = buffer.getBytes()[readIndex];
            ++readIndex;
            ++readBytes;
        }

        return readBytes;
    }

    private boolean hasBytesToRead() {
        return readIndex < buffer.size();
    }

    @Override
    public void close() throws IOException {
        if (xmlConverter != null) {
            try {
                xmlConverter.close();
            } catch (XMLStreamException e) {
                throw new RuntimeException(e);
            }
            xmlConverter = null;
        }
        input.close();
        output.close();
        super.close();
    }

    private void keepReadingFromInput() throws IOException {
        switch (readState) {
            case BEFORE_ZIP_ENTRY:
                readEntry();
                break;
            case ZIP_ENTRY_START:
                startZipEntry();
                break;
            case IN_NON_XML_ENTRY:
                readNonXmlEntry();
                break;
            case XML_ENTRY_START:
                startXmlEntry();
                break;
            case IN_XML_ENTRY:
                readXmlEntry();
                break;
        }
    }

    private void readXmlEntry() {
        try {
            if (!xmlConverter.convertNextElement()) {
                xmlConverter.close();
                xmlConverter = null;
                readState = ReadState.BEFORE_ZIP_ENTRY;
            }
        } catch (XMLStreamException e) {
            throw new RuntimeException(e);
        }
    }

    private void startXmlEntry() {
        try {
            this.xmlConverter = new OoXmlStrictConverter(OoXmlStrictConverterUtils.disableClose(input), OoXmlStrictConverterUtils.disableClose(output));
            readState = ReadState.IN_XML_ENTRY;
        } catch (XMLStreamException e) {
            throw new RuntimeException(e);
        }
    }

    private void readNonXmlEntry() throws IOException {
        int readCount = input.read(nonXmlEntryReadBuffer);
        if (-1 == readCount) {
            input.closeEntry();
            output.closeEntry();
            readState = ReadState.BEFORE_ZIP_ENTRY;
            return;
        }

        if (0 != readCount) {
            output.write(nonXmlEntryReadBuffer, 0, readCount);
            output.flush();
        }
    }

    private void startZipEntry() throws IOException {
        ZipEntry newZipEntry = new ZipEntry(zipEntry.getName());
        output.putNextEntry(newZipEntry);

        if (OoXmlStrictConverterUtils.isXml(zipEntry.getName())) {
            readState = ReadState.XML_ENTRY_START;
            return;
        }

        readState = ReadState.IN_NON_XML_ENTRY;
    }

    private void readEntry() throws IOException {
        zipEntry = input.getNextEntry();
        if (zipEntry == null) {
            readState = ReadState.END_STATE;
            output.flush();
            output.close();
            return;
        }
        readState = ReadState.ZIP_ENTRY_START;
    }

    private enum ReadState {
        BEFORE_ZIP_ENTRY,
        ZIP_ENTRY_START,
        END_STATE,
        XML_ENTRY_START,
        IN_NON_XML_ENTRY,
        IN_XML_ENTRY,
    }

    private static class ByteArrayOutputStreamExposed extends ByteArrayOutputStream {

        public byte[] getBytes() {
            return buf;
        }
    }

}
