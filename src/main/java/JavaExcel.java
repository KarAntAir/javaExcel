import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellReference;
import javax.swing.*;
import java.awt.*;
import java.io.*;
import java.nio.file.Files;
import java.nio.file.StandardCopyOption;
import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.Date;

public class JavaExcel {
    public static File TTNFile = new File("ttn.xls");
    public static File dataFile;
    private static Builder UI;

    public JavaExcel(){
        System.out.println("before UI");
        EventQueue.invokeLater(() -> {
            UI = new Builder();
            UI.prepareGUI();
            UI.showSwingUI();
            System.out.println("after UI");
        });
    }

    public static void main(String[] args) {
        System.out.println("before");
        new JavaExcel();
    }

    public static class Worker extends SwingWorker<Object, Object> {
        @Override
        protected Object doInBackground() {
            System.out.println("starting execution");
            //for (int i = 3; i < ExcelCellNumbers.rowsDataList.size(); i++)
            int startIndex = 3;
            double endIndex = ExcelCellNumbers.rowsDataList.size();
            double progressPart = 100d / endIndex;
            double currentProgress = 0d;
            try {
                for (int i = startIndex; i < endIndex; i++) {
                    System.out.println("starting " + i);
                    Row row = ExcelCellNumbers.rowsDataList.get(i);
                    File fileTTNRes = new File("./result/ttnRes"+i+".xls");
                    Files.copy(TTNFile.toPath(),fileTTNRes.toPath(), StandardCopyOption.REPLACE_EXISTING);           //копируем шаблон для дальнейшего заполнения

                    FileInputStream inputStream = new FileInputStream(fileTTNRes);

                    Workbook workbook = new HSSFWorkbook(inputStream);

                    Sheet sheet1 = workbook.getSheetAt(0); // заполняем первый лист накладной

                    fillSheet(1, sheet1, row);

                    //date stuff
                    System.out.println("getting datevalue");

                    Cell dateCell = row.getCell(12);
                    Cell FM7 = sheet1.getRow(6).getCell(168);
                    Cell FS7 = sheet1.getRow(6).getCell(174);
                    Cell FZ7 = sheet1.getRow(6).getCell(181);
                    if (dateCell.getCellType() == CellType.NUMERIC && DateUtil.isCellDateFormatted(dateCell)) {
                        System.out.println("date formatted");
                        Date date = row.getCell(12).getDateCellValue();
                        Calendar calendar = Calendar.getInstance();
                        calendar.setTime(date);

                        FM7.setCellValue(new SimpleDateFormat("dd").format(date));
                        FS7.setCellValue(new SimpleDateFormat("MM").format(date));
                        FZ7.setCellValue(new SimpleDateFormat("yyyy").format(date));
                    } else {
                        System.out.println("NOT DATE");
                        FM7.setCellValue("ОШИБКА");
                        FS7.setCellValue("ОШИБКА");
                        FZ7.setCellValue("ОШИБКА");
                    }


                    Sheet sheet2 = workbook.getSheetAt(1); // заполняем второй лист накладной

                    fillSheet(2, sheet2, row);

                    inputStream.close();

                    FileOutputStream os = new FileOutputStream(fileTTNRes);
                    workbook.write(os);
                    workbook.close();
                    os.close();
//                    inputStream = null;
//                    workbook = null;
//                    os = null;
                    currentProgress += progressPart;

                    System.out.println("current progress = "+currentProgress);
                    UI.progressBar.setValue(new Double(currentProgress).intValue());
                    System.out.println("finish " + i);

                }
                UI.setStatusText("Finish");
            } catch (Exception e) {
                System.out.println("got ex");
                System.out.println(e.getMessage());
                System.out.println(e.getStackTrace());
                System.out.println(e.getClass());
            }

            return null;
        }
    }

    public static void doWork() {
        Worker worker = new Worker();
        System.out.println("new worker");
        worker.execute();
    }

    private static void fillSheet(Integer sheetNumber, Sheet sheetToFill, Row dataRow) {
        for (CellMapping cellMapping : CellMapping.mapping.get(sheetNumber)) {
            System.out.println("current");
            System.out.println(cellMapping.getDataColumn());
            if (cellMapping.getDataColumn() == "S") {
                System.out.println("here");
            }
            System.out.println(cellMapping.getRow());
            System.out.println(cellMapping.getColumn());
            Row resultRow = sheetToFill.getRow(cellMapping.getRow());
            Cell resultCell = resultRow.getCell(CellReference.convertColStringToIndex(cellMapping.getColumn()));
            cellCopy(dataRow.getCell(CellReference.convertColStringToIndex(cellMapping.getDataColumn())), resultCell, cellMapping);
            System.out.println("copy success");
        }
    }


    public static void cellCopy(Cell cell, Cell resCell, CellMapping mapping){
        if (cell==null) {
            System.out.println("got bull cell");
            resCell.setBlank();
            System.out.println("set blank cell");
            return;
        }

        if (mapping.getReplaceStuff()) {
            if (cell.getCellType() == CellType.STRING) {
                resCell.setCellValue(CellMapping.replaceStuff(mapping, cell.getRichStringCellValue().getString()));
            } else {
                resCell.setCellValue("ОШИБКА");
            }
            return;
        }

        System.out.println("current cell type = " + cell.getCellType().name());
        switch (cell.getCellType()) {
            case STRING:
                resCell.setCellValue(cell.getRichStringCellValue().getString());
                break;
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    resCell.setCellValue(cell.getDateCellValue());
                } else {
                    System.out.println("tring to get numeric");
                    resCell.setCellValue(cell.getNumericCellValue());
                }
                break;
            case BOOLEAN:
                resCell.setCellValue(cell.getBooleanCellValue());
                break;
            case FORMULA:
                resCell.setCellValue(cell.getCellFormula());
                break;
            case BLANK:
                resCell.setBlank();
                break;
            case ERROR:
                resCell.setCellValue("ОШИБКА");
                break;
        }
    }
}
