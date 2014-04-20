package com.thundermoose.xlsx.exceptions;

/**
 * Created by tayjones on 4/20/14.
 */
public class OpenException extends RuntimeException {

  public OpenException() {
    super();
  }

  public OpenException(String msg) {
    super(msg);
  }

  public OpenException(Exception e) {
    super(e);
  }

  public OpenException(String msg, Exception e) {
    super(msg, e);
  }
}
