package org.vaadin.addons.excelexporter.configuration;

import java.util.function.Supplier;

class CacheableValue<T> {
  private T value;

  T computeIfAbsent(Supplier<T> supplier) {
    return value == null ? value = supplier.get() : value;
  }
}