package com.monitorjbl.xlsx.exceptions;

public class ParseException extends RuntimeException {

  public ParseException() {
    super();
  }

  public ParseException(String msg) {
    super(msg);
  }

  public ParseException(Exception e) {
    super(e);
  }

  public ParseException(String msg, Exception e) {
    super(msg, e);
  }
}
