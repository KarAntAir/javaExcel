import org.apache.commons.compress.utils.IOUtils;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
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
import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.Date;

public class JavaExcel {

    private File fileTTN = new File("./src/main/resources/source/ttn.xls");

    public static void main(String[] args) throws IOException, URISyntaxException {
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

            Row row1 = sheet1.getRow(8);
            Cell V9 = row1.getCell(21);
            cellCopy(row.getCell(13),V9);

            Cell FM9 = row1.getCell(168);
            cellCopy(row.getCell(13),FM9);

            Cell FM6 = sheet1.getRow(5).getCell(168);
            cellCopy(row.getCell(11),FM6);


            Date date = row.getCell(12).getDateCellValue();
            Calendar calendar = Calendar.getInstance();
            calendar.setTime(date);

            Cell FM7 = sheet1.getRow(6).getCell(168);
            Cell FS7 = sheet1.getRow(6).getCell(174);
            Cell FZ7 = sheet1.getRow(6).getCell(181);
            FM7.setCellValue(new SimpleDateFormat("dd").format(date));
            FS7.setCellValue(new SimpleDateFormat("MM").format(date));
            FZ7.setCellValue(new SimpleDateFormat("yyyy").format(date));


            Cell AS18 = sheet1.getRow(17).getCell(44);
            cellCopy(row.getCell(7),AS18);

            Cell BU18 = sheet1.getRow(17).getCell(72);
            cellCopy(row.getCell(4),BU18);

            Cell ED44 = sheet1.getRow(43).getCell(133);
            cellCopy(row.getCell(15),ED44);

            Cell FI44 = sheet1.getRow(43).getCell(164);
            cellCopy(row.getCell(16),FI44);

            Cell BR42 = sheet1.getRow(41).getCell(69);
            cellCopy(row.getCell(21),BR42);

            Cell BB42 = sheet1.getRow(41).getCell(CellReference.convertColStringToIndex("BB"));
            setPict(workbook,sheet1,BB42);


            Sheet sheet2 = workbook.getSheetAt(1); // заполняем второй лист накладной

            Cell N4 = sheet2.getRow(3).getCell(13);
            cellCopy(row.getCell(13),N4);

            Cell CO4 = sheet2.getRow(3).getCell(92);
            cellCopy(row.getCell(9),CO4);

            Cell EL4 = sheet2.getRow(3).getCell(141);
            cellCopy(row.getCell(10),EL4);

            Cell Q14 = sheet2.getRow(13).getCell(16);
            cellCopy(row.getCell(18),Q14);

            Cell FD34 = sheet2.getRow(33).getCell(159);
            cellCopy(row.getCell(22),FD34);

            Cell FD36 = sheet2.getRow(35).getCell(159);
            cellCopy(row.getCell(23),FD36);
            //System.out.println(CellReference.convertColStringToIndex("FD"));



            inputStream.close();

            FileOutputStream os = new FileOutputStream(fileTTNRes);
            workbook.write(os);
            workbook.close();
            os.close();
        }

    }


    public void cellCopy(Cell cell, Cell resCell){
        if (cell==null) {
            resCell.setBlank();
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
                new File("./src/main/resources/source/pictures/3.png"));

        BufferedImage SubImg = rotate(originalImg);
        ByteArrayOutputStream baos = new ByteArrayOutputStream();
        ImageIO.write(SubImg, "png", baos);
        byte[] bytes = baos.toByteArray();


        // Добавляем изображение в документ
        int pictureIdx = workbook.addPicture(bytes, Workbook.PICTURE_TYPE_PNG);
        CreationHelper helper = workbook.getCreationHelper();
        Drawing drawing = sheet.createDrawingPatriarch();

        ClientAnchor pechatAnc = helper.createClientAnchor();

        pechatAnc.setCol1(90);
        pechatAnc.setCol2(140);
        pechatAnc.setRow1(28);
        pechatAnc.setRow2(48);

        drawing.createPicture(pechatAnc, pictureIdx);

    }

    public static BufferedImage rotate(BufferedImage img) {

        int width = img.getWidth();
        int height = img.getHeight();

        BufferedImage newImage = new BufferedImage(
                img.getWidth(), img.getHeight(), img.getType());

        Graphics2D g2 = newImage.createGraphics();

        g2.rotate(Math.toRadians(45), width / 2, height / 2);
        g2.drawImage(img, null, 0, 0);

        return newImage;
    }
}
