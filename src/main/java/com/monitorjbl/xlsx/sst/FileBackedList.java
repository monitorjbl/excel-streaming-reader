package com.monitorjbl.xlsx.sst;

import com.fasterxml.jackson.annotation.JsonInclude.Include;
import com.fasterxml.jackson.databind.ObjectMapper;

import java.io.File;
import java.io.IOException;
import java.io.RandomAccessFile;
import java.nio.ByteBuffer;
import java.nio.channels.FileChannel;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

/**
 * File-backed list-like class. Allows addition of arbitrary
 * numbers of array entries (serialized to JSON) in a binary
 * packed file. Reading of entries is done with an NIO
 * channel that seeks to the entry in the file.
 * <p>
 * File entry format:
 * <ul>
 * <li>4 bytes: length of entry</li>
 * <li><i>length</i> bytes: JSON string containing the entry data</li>
 * </ul>
 * <p>
 * Pointers to the offset of each entry are kept in a {@code List<Long>}.
 * The values loaded from the the file are cached up to a maximum of
 * {@code cacheSize}. Items are evicted from the cache with an LRU algorithm.
 */
public class FileBackedList<T> implements AutoCloseable {
  private final static ObjectMapper mapper;

  static {
    mapper = new ObjectMapper().setSerializationInclusion(Include.NON_NULL);
  }

  private final Class<T> type;
  private final List<Long> pointers = new ArrayList<>();
  private final RandomAccessFile raf;
  private final FileChannel channel;
  private final Map<Integer, T> cache;

  private long filesize;

  public FileBackedList(Class<T> type, File file, final int cacheSize) throws IOException {
    this.type = type;
    this.raf = new RandomAccessFile(file, "rw");
    this.channel = raf.getChannel();
    this.filesize = raf.length();
    this.cache = new LinkedHashMap<Integer, T>(cacheSize, 0.75f, true) {
      public boolean removeEldestEntry(Map.Entry eldest) {
        return size() > cacheSize;
      }
    };
  }

  public void add(T obj) {
    try {
      writeToFile(obj);
    } catch(IOException e) {
      throw new RuntimeException(e);
    }
  }

  public T getAt(int index) {
    if(cache.containsKey(index)) {
      return cache.get(index);
    }

    try {
      T val = readFromFile(pointers.get(index));
      cache.put(index, val);
      return val;
    } catch(IOException e) {
      throw new RuntimeException(e);
    }
  }

  private void writeToFile(T obj) throws IOException {
    synchronized (channel) {
      ByteBuffer bytes = ByteBuffer.wrap(mapper.writeValueAsBytes(obj));
      ByteBuffer length = ByteBuffer.allocate(4).putInt(bytes.array().length);

      channel.position(filesize);
      pointers.add(channel.position());
      length.flip();
      channel.write(length);
      channel.write(bytes);

      filesize += 4 + bytes.array().length;
    }
  }

  private T readFromFile(long pointer) throws IOException {
    synchronized (channel) {
      FileChannel fc = channel.position(pointer);

      //get length of entry
      ByteBuffer buffer = ByteBuffer.wrap(new byte[4]);
      fc.read(buffer);
      buffer.flip();
      int length = buffer.getInt();

      //read entry
      buffer = ByteBuffer.wrap(new byte[length]);
      fc.read(buffer);
      buffer.flip();

      return mapper.readValue(buffer.array(), type);
    }
  }

  @Override
  public void close() {
    try {
      raf.close();
    } catch(IOException e) {
      throw new RuntimeException(e);
    }
  }
}
