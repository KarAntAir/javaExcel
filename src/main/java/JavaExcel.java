import org.apache.commons.compress.utils.IOUtils;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellAddress;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFClientAnchor;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


import javax.imageio.ImageIO;
import java.awt.*;
import java.awt.image.BufferedImage;
import java.io.*;
import java.net.URISyntaxException;
import java.nio.file.Files;
import java.nio.file.StandardCopyOption;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Date;
import java.util.HashMap;

public class JavaExcel {


    private File fileTTN = new File("ttn.xls");


    public static void main(String[] args) throws IOException, URISyntaxException, InvalidFormatException {
        ExcelCellNumbers.fillRowsDataList();
        new JavaExcel().doSmth();
    }

    public void doSmth() throws IOException, InvalidFormatException {
        //for (int i = 3; i < ExcelCellNumbers.rowsDataList.size(); i++)
        for (int i = 3; i < 7; i++) {
            Row row = ExcelCellNumbers.rowsDataList.get(i);
            File fileTTNRes = new File("ttnRes"+i+".xls");
            Files.copy(fileTTN.toPath(),fileTTNRes.toPath(), StandardCopyOption.REPLACE_EXISTING);           //копируем шаблон для дальнейшего заполнения

            FileInputStream inputStream = new FileInputStream(fileTTNRes);

            Workbook workbook = new HSSFWorkbook(inputStream);


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


            Cell BD39 = sheet1.getRow(41).getCell(CellReference.convertColStringToIndex("BD"));
            setPict(workbook,sheet1,BD39);

            Sheet sheet2 = workbook.getSheetAt(1); // заполняем второй лист накладной

            fillSheet(2, sheet2, row);



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
        BufferedImage originalImg = ImageIO.read(
                new File("3.png"));

        BufferedImage SubImg = rotate(originalImg);
        ByteArrayOutputStream baos = new ByteArrayOutputStream();
        ImageIO.write(SubImg, "png", baos);
        byte[] bytes = baos.toByteArray();



        // Добавляем изображение в документ
        int pictureIdx = workbook.addPicture(bytes, Workbook.PICTURE_TYPE_PNG);
        CreationHelper helper = workbook.getCreationHelper();
        Drawing drawing = sheet.createDrawingPatriarch();

        // Создаем якорь для изображения
        ClientAnchor pechatAnc = helper.createClientAnchor();


        pechatAnc.setCol1(45);
        pechatAnc.setCol2(95);
        pechatAnc.setRow1(26);
        pechatAnc.setRow2(46);


        drawing.createPicture(pechatAnc, pictureIdx);

    }

    public static BufferedImage rotate(BufferedImage img) {

        int width = img.getWidth();
        int height = img.getHeight();

        BufferedImage newImage = new BufferedImage(
                img.getWidth(), img.getHeight(), img.getType());

        Graphics2D g2 = newImage.createGraphics();

        int rotationAngle = (int) (Math.random()  * 61) - 30;
        g2.rotate(Math.toRadians(rotationAngle), width / 2, height / 2);
        g2.drawImage(img, null, 0, 0);


        return newImage;
    }
}
