import org.apache.poi.ss.formula.functions.Column;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellAddress;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class ExcelCellNumbers {

//    public static String [] list1 = {"V9", "FM9","FM6","FM7","FS7","FZ7","AS18","BU18","ED44","FI44",
//            "BR42","BD39","BB42","EV44","EJ39"};
//    public static String [] list2 = {"N4","CO4","EL4","Q14","FD34","FD36"};

    public static List<Row> rowsDataList = new ArrayList<>();



    public static void fillRowsDataList() throws IOException {
        FileInputStream fileInputStream = new FileInputStream(new File("./src/main/resources/source/Песок с указанием карьеров СТГ.xlsx"));
        Workbook workbook = WorkbookFactory.create(fileInputStream);
        Sheet sheet = workbook.getSheetAt(0);
        for (Row row: sheet){
            rowsDataList.add(row);
        }
        fileInputStream.close();
        workbook.close();
    }

}


