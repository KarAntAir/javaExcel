import org.apache.commons.compress.utils.IOUtils;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellAddress;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFClientAnchor;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


import java.io.*;
import java.nio.file.Files;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Date;
import java.util.HashMap;

public class JavaExcel {
    private File fileTTN = new File("./src/main/resources/source/ttn.xls");;



    public static void main(String[] args) throws IOException {
        ExcelCellNumbers.fillRowsDataList();
        new JavaExcel().doSmth();
    }

    public void doSmth() throws IOException {
        //for (int i = 3; i < ExcelCellNumbers.rowsDataList.size(); i++)
        for (int i = 3; i < 7; i++) {
            Row row = ExcelCellNumbers.rowsDataList.get(i);
            File fileTTNRes = new File("./src/main/resources/result/ttnRes"+i+".xls");
            Files.copy(fileTTN.toPath(),fileTTNRes.toPath());           //копируем шаблон для дальнейшего заполнения

            FileInputStream inputStream = new FileInputStream(fileTTNRes);
            Workbook workbook = WorkbookFactory.create(inputStream);

            Sheet sheet1 = workbook.getSheetAt(0); // заполняем первый лист накладной

            fillSheet(1, sheet1, row);

            //date stuff
            Date date = row.getCell(12).getDateCellValue();
            Calendar calendar = Calendar.getInstance();
            calendar.setTime(date);

            Cell FM7 = sheet1.getRow(6).getCell(168);
            Cell FS7 = sheet1.getRow(6).getCell(174);
            Cell FZ7 = sheet1.getRow(6).getCell(181);
            FM7.setCellValue(new SimpleDateFormat("dd").format(date));
            FS7.setCellValue(new SimpleDateFormat("MM").format(date));
            FZ7.setCellValue(new SimpleDateFormat("yyyy").format(date));

//            Cell BB42 = sheet1.getRow(41).getCell(CellReference.convertColStringToIndex("BB"));
//            setPict(workbook,sheet1,BB42);


            Sheet sheet2 = workbook.getSheetAt(1); // заполняем второй лист накладной

            fillSheet(2, sheet2, row);

            //System.out.println(CellReference.convertColStringToIndex("FD"));


            inputStream.close();

            FileOutputStream os = new FileOutputStream(fileTTNRes);
            workbook.write(os);
            workbook.close();
            os.close();
        }

    }

    private void fillSheet(Integer sheetNumber, Sheet sheetToFill, Row dataRow) {
        for (CellMapping cellMapping : CellMapping.mapping.get(sheetNumber)) {
            Row resultRow = sheetToFill.getRow(cellMapping.getRow());
            Cell resultCell = resultRow.getCell(CellReference.convertColStringToIndex(cellMapping.getColumn()));
            cellCopy(dataRow.getCell(CellReference.convertColStringToIndex(cellMapping.getDataRow())), resultCell, cellMapping);
        }
    }


    public void cellCopy(Cell cell, Cell resCell, CellMapping mapping){
        if (cell==null) {
            resCell.setBlank();
            return;
        }

        if (mapping.getReplaceStuff()) {
            resCell.setCellValue(CellMapping.replaceStuff(mapping, cell.getRichStringCellValue().getString()));
            return;
        }

        switch (cell.getCellType()) {
            case STRING:
                resCell.setCellValue(cell.getRichStringCellValue().getString());
                break;
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    resCell.setCellValue(cell.getDateCellValue());
                } else {
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
        }
    }


    public void setPict(Workbook workbook, Sheet sheet, Cell cell) throws IOException {

//        anchor.setCol1(cell.getColumnIndex()); // Sets the column (0 based) of the first cell.
//        anchor.setCol2(cell.getColumnIndex()+1); // Sets the column (0 based) of the Second cell.
//        anchor.setRow1(cell.getRowIndex()); // Sets the row (0 based) of the first cell.
//        anchor.setRow2(cell.getRowIndex()+1); // Sets the row (0 based) of the Second cell.


        // Загружаем изображение из файла
        FileInputStream imageFile = new FileInputStream("./src/main/resources/source/pictures/2.jpg");
        byte[] imageBytes = IOUtils.toByteArray(imageFile);
        imageFile.close();

        // Добавляем изображение в документ
        int pictureIdx = workbook.addPicture(imageBytes, Workbook.PICTURE_TYPE_JPEG);
        CreationHelper helper = workbook.getCreationHelper();
        Drawing drawing = sheet.createDrawingPatriarch();

        // Создаем якорь для изображения
        ClientAnchor anchor = helper.createClientAnchor();
        anchor.setCol1(cell.getColumnIndex()); // Установите номер столбца, в котором должно быть изображение
        anchor.setCol2(cell.getColumnIndex()+1);
        anchor.setRow1(cell.getRowIndex()); // Установите номер строки, в которой должно быть изображение
        anchor.setRow2(cell.getRowIndex()+1);

        // Создаем объект картинки
        Picture picture = drawing.createPicture(anchor, pictureIdx);
        //picture.resize(); // Масштабируем изображение, чтобы оно подходило к размерам ячейки
    }
}
