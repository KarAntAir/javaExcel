import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.ArrayList;
import java.util.List;

public class ExcelCellNumbers {

    public static List<Row> rowsDataList = new ArrayList<>();



    public static void fillRowsDataList() throws IOException, InvalidFormatException {
//        OPCPackage pkg = OPCPackage.open(new File("Песок с указанием карьеров СТГ.xlsx"));
        System.out.println("USING DATA FILE AS " +JavaExcel.dataFile.getAbsolutePath());
        OPCPackage pkg = OPCPackage.open(new FileInputStream(JavaExcel.dataFile));
        XSSFWorkbook workbook = new XSSFWorkbook(pkg);
        Sheet sheet = workbook.getSheetAt(0);
        for (Row row: sheet){
            rowsDataList.add(row);
        }
        pkg.close();
    }

}


