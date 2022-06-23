package org.example;


import com.opencsv.*;
import com.opencsv.bean.CsvToBeanBuilder;
import org.apache.commons.io.FilenameUtils;
import org.apache.commons.io.input.BOMInputStream;

import java.io.*;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.List;

/**
 * Hello world!
 *
 */
public class App {

    public static void main( String[] args ) {

        try {
            ExcelHandler excelHandler = new ExcelHandler();
            //0：文件路径， 1：开始位置 2：结束为止  3：解析类型（1 COC01~XX，2 指定单元格） 4：指定单元格（可选，由3决定）
            String productNumber = args[0];

            // 从CSV读Rule
            List<ProductRule> productRules;
            RFC4180Parser rfc4180Parser = new RFC4180ParserBuilder().build();
            try(Reader reader = new InputStreamReader(new BOMInputStream(Files.newInputStream(Paths.get("ProductRule.csv"))));
                CSVReader csvReader = new CSVReaderBuilder(reader).withCSVParser(rfc4180Parser).build();) {
                productRules = new ArrayList<>(new CsvToBeanBuilder<ProductRule>(csvReader).withType(ProductRule.class).withQuoteChar(CSVWriter.NO_QUOTE_CHARACTER).build().parse());
            }

            ProductRule productRule = productRules.stream().filter(it->it.getProductNumber().trim().contentEquals(productNumber.trim())).findFirst().orElseThrow(()->new RuntimeException("未找到"+productNumber+"的料号规则"));

            File userDir = new File(System.getProperty("user.dir"));
            File[] excelFiles = userDir.listFiles(file ->
                    file.isFile()
                            && (FilenameUtils.getExtension(file.getName()).equals("xls") || FilenameUtils.getExtension(file.getName()).equals("xlsx")));

            if (productRule.getParseType() == 2){
                excelHandler.handlerExcelFiles(excelFiles, productRule.getChuckSize(), productRule.getParseType(), productRule.getCellAddress());
            }else {
                excelHandler.handlerExcelFiles(excelFiles, productRule.getChuckSize(), productRule.getParseType(), null);
            }
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }
}
