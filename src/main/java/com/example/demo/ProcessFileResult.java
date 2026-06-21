package com.example.demo;

import java.util.Map;
import java.util.Set;

public class ProcessFileResult {

    Map<String, Set<Integer>> excludedColsPerSheet;

    Set<String> collectValuesFromFile;

    public Map<String, Set<Integer>> getExcludedColsPerSheet() {
        return excludedColsPerSheet;
    }

    public void setExcludedColsPerSheet(Map<String, Set<Integer>> excludedColsPerSheet) {
        this.excludedColsPerSheet = excludedColsPerSheet;
    }

    public Set<String> getCollectValuesFromFile() {
        return collectValuesFromFile;
    }

    public void setCollectValuesFromFile(Set<String> collectValuesFromFile) {
        this.collectValuesFromFile = collectValuesFromFile;
    }
}
