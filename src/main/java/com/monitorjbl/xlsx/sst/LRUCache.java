package com.monitorjbl.xlsx.sst;

import java.util.Iterator;
import java.util.LinkedHashMap;

class LRUCache  {

    private long sizeBytes;
    private final long capacityBytes;
    private final LinkedHashMap<Integer, String> map = new LinkedHashMap<>();

    LRUCache(long capacityBytes) {
        this.capacityBytes = capacityBytes;
    }

    String getIfPresent(int key) {
        String s = map.get(key);
        if (s != null) {
            map.remove(key);
            map.put(key, s);
        }
        return s;
    }

    void store(int key, String val) {
        long valSize = strSize(val);
        if (valSize > capacityBytes)
            throw new RuntimeException("Insufficient cache space.");
        Iterator<String> it = map.values().iterator();
        while (valSize + sizeBytes > capacityBytes) {
            String s = it.next();
            sizeBytes -= strSize(s);
            it.remove();
        }
        map.put(key, val);
        sizeBytes += valSize;
    }

//  just an estimation
    private static long strSize(String str) {
        long size = Integer.BYTES; // hashCode
        size += Character.BYTES * str.length(); // characters
        return size;
    }

}
