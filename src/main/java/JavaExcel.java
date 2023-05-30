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
    private JPanel templateSelectorPanel;
    private JPanel templateSelectorPathPanel;
    private JPanel dataSelectorPanel;
    private JPanel dataSelectorPathPanel;
    private JPanel statusBarPanel;
    private JPanel actionPanel;
    private static JProgressBar progressBar;

    public JavaExcel(){
        prepareGUI();
    }

    private void prepareGUI(){
        mainFrame = new JFrame("SWING TRY");
        mainFrame.setSize(400,400);
        mainFrame.setLayout(new GridLayout(6, 1));

        mainFrame.addWindowListener(new WindowAdapter() {
            public void windowClosing(WindowEvent windowEvent){
                System.exit(0);
            }
        });
        templateSelectorPanel = new JPanel();
        templateSelectorPathPanel = new JPanel();
        dataSelectorPanel = new JPanel();
        dataSelectorPathPanel = new JPanel();
        statusBarPanel = new JPanel();
        actionPanel = new JPanel();

        mainFrame.add(templateSelectorPanel);
        mainFrame.add(templateSelectorPathPanel);
        mainFrame.add(dataSelectorPanel);
        mainFrame.add(dataSelectorPathPanel);
        mainFrame.add(statusBarPanel);
        mainFrame.add(actionPanel);
        mainFrame.setVisible(true);
    }

    private void showSwingUI(){
        JPanel templateSelectorPan = new JPanel();
        JLabel templateSelectorText = new JLabel("Выберете файл шаблона");
        JButton templateSelectorButton = new JButton("Выбрать шаблон");
        templateSelectorPan.add(templateSelectorText);
        templateSelectorPan.add(templateSelectorButton);

        JPanel templatePathPan = new JPanel();
        JLabel templateResultText = new JLabel("Файл не выбран");
        templatePathPan.add(templateResultText);
        templateSelectorPathPanel.add(templatePathPan);

        ActionListener dataListener = new DataListener(templateResultText);
        templateSelectorButton.addActionListener(dataListener);
        templateSelectorPanel.add(templateSelectorPan);

        JPanel dataSelectorPan = new JPanel();
        JLabel dataSelectorText = new JLabel("Выберете файл данных");
        JButton dataSelectorButton = new JButton("Выбрать файл данных");
        dataSelectorPan.add(dataSelectorText);
        dataSelectorPan.add(dataSelectorButton);

        JPanel dataPathPan = new JPanel();
        JLabel dataResultText = new JLabel("Файл не выбран");
        dataPathPan.add(dataResultText);
        dataSelectorPathPanel.add(dataPathPan);

        ActionListener templateL = new DataListener(dataResultText);
        dataSelectorButton.addActionListener(templateL);
        dataSelectorPanel.add(dataSelectorPan);

        JPanel barPan = new JPanel();
        progressBar = new JProgressBar();
        progressBar.setValue(0);
        progressBar.setStringPainted(true);
        barPan.add(progressBar);
        statusBarPanel.add(barPan);

        JPanel actionPan = new JPanel();
        JButton actionButton = new JButton("Run");
        actionPan.add(actionButton);
        ActionListener runListener = new RunListener();
        actionButton.addActionListener(runListener);
        actionPanel.add(actionPan);

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
                if (com.equals("Выбрать файл данных")) {
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
}
