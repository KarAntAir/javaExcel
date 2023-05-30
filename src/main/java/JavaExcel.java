import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellReference;


import javax.imageio.ImageIO;
import javax.swing.*;
import java.awt.*;
import java.awt.event.*;
import java.awt.image.BufferedImage;
import java.io.*;
import java.nio.file.Files;
import java.nio.file.StandardCopyOption;
import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.Date;

public class JavaExcel {

//    private File fileTTN = new File("ttn.xls");
    private static File TTNFile;
    public static File dataFile;

    private JFrame mainFrame;
    private JPanel firstPanel;
    private JPanel secondPanel;
    private JPanel barPanel;
    private JPanel actionPanel;
    private static JProgressBar progressBar;

    public JavaExcel(){
        prepareGUI();
    }

    private void prepareGUI(){
        mainFrame = new JFrame("SWING TRY");
        mainFrame.setSize(400,400);
        mainFrame.setLayout(new GridLayout(4, 1));

        mainFrame.addWindowListener(new WindowAdapter() {
            public void windowClosing(WindowEvent windowEvent){
                System.exit(0);
            }
        });
        firstPanel = new JPanel();
        secondPanel = new JPanel();
        barPanel = new JPanel();
        actionPanel = new JPanel();

        mainFrame.add(firstPanel);
        mainFrame.add(secondPanel);
        mainFrame.add(barPanel);
        mainFrame.add(actionPanel);
        mainFrame.setVisible(true);
    }

    private void showSwingUI(){
        JPanel fPanel = new JPanel();
        JLabel fLabel = new JLabel("select template file");
        JButton fButton = new JButton("select ttn file");
        JLabel fLabel2 = new JLabel("no file selected");
        fPanel.add(fLabel);
        fPanel.add(fButton);
        fPanel.add(fLabel2);
        ActionListener dataL = new DataListener(fLabel2);
        fButton.addActionListener(dataL);
        firstPanel.add(fPanel);

        JPanel panel2 = new JPanel();
        JLabel sLabel = new JLabel("select data file");
        JButton sButton = new JButton("select data file");
        JLabel sLabel2 = new JLabel("no file selected");
        panel2.add(sLabel);
        panel2.add(sButton);
        panel2.add(sLabel2);
        ActionListener templateL = new DataListener(sLabel2);
        sButton.addActionListener(templateL);
        secondPanel.add(panel2);

        JPanel barPan = new JPanel();
        progressBar = new JProgressBar();
        progressBar.setValue(0);
        progressBar.setStringPainted(true);
        barPan.add(progressBar);
        barPanel.add(barPan);

        JPanel panel3 = new JPanel();
        JButton actionButton = new JButton("Run");
        panel3.add(actionButton);
        ActionListener runListener = new RunListener();
        actionButton.addActionListener(runListener);
        actionPanel.add(panel3);

        mainFrame.setVisible(true);
    }

    static class RunListener implements ActionListener {
        @Override
        public void actionPerformed(ActionEvent e) {
            try {
                ExcelCellNumbers.fillRowsDataList();
                JavaExcel.doSmth();
            } catch (IOException | InvalidFormatException ex) {
                throw new RuntimeException(ex);
            }
        }
    }

    static class DataListener implements ActionListener {

        private JLabel label;
        DataListener(JLabel l) {
            this.label = l;
        }
        @Override
        public void actionPerformed(ActionEvent e) {
            File workingDirectory = new File(System.getProperty("user.dir"));
            JFileChooser j = new JFileChooser();
            j.setCurrentDirectory(workingDirectory);
            int result = j.showOpenDialog(null);
            if (result == JFileChooser.APPROVE_OPTION) {
                label.setText(j.getSelectedFile().getAbsolutePath());
                String com = e.getActionCommand();
                if (com.equals("select data file")) {
                    JavaExcel.dataFile = new File(j.getSelectedFile().getAbsolutePath());
                    System.out.println("setting data file as " + j.getSelectedFile().getAbsolutePath());
                } else {
                    JavaExcel.TTNFile = new File(j.getSelectedFile().getAbsolutePath());
                    System.out.println("setting ttn file as " + j.getSelectedFile().getAbsolutePath());

                }
            } else {
                label.setText("cancel");
            }
        }
    }

    public static void main(String[] args) {
//        ExcelCellNumbers.fillRowsDataList();
//        new JavaExcel().doSmth();

        JavaExcel javaExcel = new JavaExcel();
        javaExcel.showSwingUI();
    }

    public static void doSmth() throws IOException {
        //for (int i = 3; i < ExcelCellNumbers.rowsDataList.size(); i++)
        int startIndex = 3;
        int endIndex = 7;
        int progressPart = (100 / (endIndex - startIndex));
        int currentProgress = 0;
        for (int i = 3; i < endIndex; i++) {
            Row row = ExcelCellNumbers.rowsDataList.get(i);
            File fileTTNRes = new File("ttnRes"+i+".xls");
//            Files.copy(fileTTN.toPath(),fileTTNRes.toPath(), StandardCopyOption.REPLACE_EXISTING);           //копируем шаблон для дальнейшего заполнения
            Files.copy(TTNFile.toPath(),fileTTNRes.toPath(), StandardCopyOption.REPLACE_EXISTING);           //копируем шаблон для дальнейшего заполнения

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
            //setPict(workbook,sheet1,BD39);

            Sheet sheet2 = workbook.getSheetAt(1); // заполняем второй лист накладной

            fillSheet(2, sheet2, row);



            inputStream.close();

            FileOutputStream os = new FileOutputStream(fileTTNRes);
            workbook.write(os);
            workbook.close();
            os.close();
            currentProgress += progressPart;
            progressBar.setValue(currentProgress);
        }

    }

    private static void fillSheet(Integer sheetNumber, Sheet sheetToFill, Row dataRow) {
        for (CellMapping cellMapping : CellMapping.mapping.get(sheetNumber)) {
            Row resultRow = sheetToFill.getRow(cellMapping.getRow());
            Cell resultCell = resultRow.getCell(CellReference.convertColStringToIndex(cellMapping.getColumn()));
            cellCopy(dataRow.getCell(CellReference.convertColStringToIndex(cellMapping.getDataRow())), resultCell, cellMapping);
        }
    }


    public static void cellCopy(Cell cell, Cell resCell, CellMapping mapping){
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


    public static void setPict(Workbook workbook, Sheet sheet, Cell cell) throws IOException {
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
