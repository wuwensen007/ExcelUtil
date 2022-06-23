package org.example;

import com.opencsv.bean.CsvBindByName;

public class ProductRule {

    @CsvBindByName(required = true)
    private String ProductNumber;

    @CsvBindByName(required = true)
    private int ChuckSize;

    @CsvBindByName(required = true)
    private int ParseType;

    @CsvBindByName()
    private String CellAddress;

    public String getProductNumber() {
        return ProductNumber;
    }

    public void setProductNumber(String productNumber) {
        ProductNumber = productNumber;
    }

    public int getChuckSize() {
        return ChuckSize;
    }

    public void setChuckSize(int chuckSize) {
        ChuckSize = chuckSize;
    }

    public int getParseType() {
        return ParseType;
    }

    public void setParseType(int parseType) {
        ParseType = parseType;
    }

    public String getCellAddress() {
        return CellAddress;
    }

    public void setCellAddress(String cellAddress) {
        CellAddress = cellAddress;
    }
}
