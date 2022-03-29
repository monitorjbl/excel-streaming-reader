package com.github.pjfanning.xlsx;

import java.io.Closeable;
import java.util.Iterator;

/**
 * An iterator that should be closed after use
 *
 * @param <T> type param for iterator
 */
public interface CloseableIterator<T> extends Iterator<T>, Closeable {

}
