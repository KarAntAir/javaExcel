import com.ibm.icu.text.RuleBasedNumberFormat;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.util.CellReference;

import javax.swing.*;
import java.awt.*;
import java.io.*;

public class JavaExcel {
    public static File dataFile;
    public static Builder UI;

    public JavaExcel() {
        EventQueue.invokeLater(() -> {
            UI = new Builder();
            UI.prepareGUI();
            UI.showSwingUI();
        });
    }

    public static void main(String[] args) {
//        String s1 = "ООО \"КОРОНА РУС\" Москва, ул. Каспийская, д. 22, корпус 1, стр. 5, этаж 5, пом. 9, к. 17А, офис 86 телефон: +7\u00A0(925)\u00A0641-85-57 р/с 40702810801840000325 АО \"АЛЬФА-БАНК\" к/с 30101810200000000593, БИК 044525593";
//        String s2 = "ООО \"КОРОНА РУС\" Москва, ул. Каспийская, д. 22, корпус 1, стр. 5, этаж 5, пом. 9, к. 17А, офис 86 телефон: +7 (925) 641-85-57 р/с 40702810801840000325 АО \"АЛЬФА-БАНК\" к/с 30101810200000000593, БИК 044525593";
//        System.out.println(s1.length());
//        System.out.println(s2.length());
//        String s = "ООО \"КОРОНА РУС\" Москва, ул. Каспийская, д. 22, корпус 1, стр. 5, этаж 5, пом. 9, к. 17А, офис 86 телефон: +7 (925) 641-85-57 р/с 40702810801840000325 АО \"АЛЬФА-БАНК\" к/с 30101810200000000593, БИК 044525593ООО \"КОРОНА РУС\" Москва, ул. Каспийская, д. 22, корпус 1, стр. 5, этаж 5, пом. 9, к. 17А, офис 86 телефон: +7 (925) 641-85-57 р/с 40702810801840000325 АО \"АЛЬФА-БАНК\" к/с 30101810200000000593, БИК 044525593";
//        new JavaExcel();
//        System.out.println(s);
//        System.out.println(s1);

        RuleBasedNumberFormat nf = new RuleBasedNumberFormat(Locale.forLanguageTag("ru"),
                RuleBasedNumberFormat.SPELLOUT);
        System.out.println(nf.format(1234567.0));
//        try {
//            PDFConverter.saveAsPdf2();
//
//        } catch (Exception e) {
//            throw new RuntimeException(e);
//        }
    }

    public static void doWork() {
        Worker worker = new Worker();
        worker.execute();
    }

    public static class Worker extends SwingWorker<Object, Object> {
        private static HashMap<String, String> errors = new HashMap<String, String>();
        private static File logFile;

        private Boolean isValidName(String text) {
            final Character[] INVALID_WINDOWS_SPECIFIC_CHARS = {':', '\\', '/', '"', '*', '<', '>', '?', '|'};
            return Arrays.stream(INVALID_WINDOWS_SPECIFIC_CHARS)
                    .noneMatch(ch -> text.contains(ch.toString()));
        }

        @Override
        protected Object doInBackground() {
            int startIndex = 3;
            double endIndex = 5;//ExcelCellNumbers.rowsDataList.size();
            double progressPart = 100d / endIndex;
            double currentProgress = 0d;
            String timeStamp = new SimpleDateFormat("yyyyMMdd-HHmmss").format(Calendar.getInstance().getTime());
            String newDirectory = "result-" + timeStamp;
            File dir = new File(newDirectory);
            if (!dir.exists()) {
                Boolean b = dir.mkdir();
            }
            createLogFile(newDirectory);
            try {
                for (int i = startIndex; i < endIndex; i++) {
                    Row row = ExcelCellNumbers.rowsDataList.get(i);

                    Cell ttnNumber = row.getCell(CellReference.convertColStringToIndex("L"));
                    String fileNameTTNPart = getFileNameTTNPart(i, ttnNumber);

                    File fileTTNRes = new File(newDirectory + File.separatorChar + fileNameTTNPart + ".xls");
                    Files.copy(TTNFile.toPath(), fileTTNRes.toPath(), StandardCopyOption.REPLACE_EXISTING);           //копируем шаблон для дальнейшего заполнения

                    FileInputStream inputStream = new FileInputStream(fileTTNRes);

                    Workbook workbook = new HSSFWorkbook(inputStream);

                    Sheet sheet1 = workbook.getSheetAt(0); // заполняем первый лист накладной

                    fillSheet(1, sheet1, row);

                    //date stuff

                    Cell dateCell = row.getCell(12);
                    Cell FM7 = sheet1.getRow(6).getCell(168);
                    Cell FS7 = sheet1.getRow(6).getCell(174);
                    Cell FZ7 = sheet1.getRow(6).getCell(181);
                    if (dateCell.getCellType() == CellType.NUMERIC && DateUtil.isCellDateFormatted(dateCell)) {
                        Date date = row.getCell(12).getDateCellValue();
                        Calendar calendar = Calendar.getInstance();
                        calendar.setTime(date);

                        FM7.setCellValue(new SimpleDateFormat("dd").format(date));
                        FS7.setCellValue(new SimpleDateFormat("MM").format(date));
                        FZ7.setCellValue(new SimpleDateFormat("yyyy").format(date));
                    } else {
                        errors.put("Ошибка даты", "в данных не найдена дата");
                        FM7.setCellValue("ОШИБКА");
                        FS7.setCellValue("ОШИБКА");
                        FZ7.setCellValue("ОШИБКА");
                    }


                    Sheet sheet2 = workbook.getSheetAt(1); // заполняем второй лист накладной

                    fillSheet(2, sheet2, row);

                    FormulaEvaluator formulaEvaluator = workbook.getCreationHelper().createFormulaEvaluator();
                    formulaEvaluator.evaluateAll();

                    workbook.setPrintArea(
                            0, //sheet num
                            0, //start column
                            186, //end column
                            0, //start row
                            44 //end row
                    );
                    workbook.setPrintArea(
                            1, //sheet num
                            0, //start column
                            189, //end column
                            0, //start row
                            60 //end row
                    );

                    sheet1.getPrintSetup().setPaperSize(PrintSetup.A4_PAPERSIZE);
                    sheet2.getPrintSetup().setPaperSize(PrintSetup.A4_PAPERSIZE);
                    sheet1.setFitToPage(true);
                    sheet2.setFitToPage(true);
                    sheet1.getPrintSetup().setFitWidth((short)1);
                    sheet1.getPrintSetup().setFitHeight((short)1);
                    sheet2.getPrintSetup().setFitWidth((short)1);
                    sheet2.getPrintSetup().setFitHeight((short)1);




                    inputStream.close();

                    FileOutputStream os = new FileOutputStream(fileTTNRes);
                    workbook.write(os);
                    workbook.close();
                    os.close();
                    fillLogFile(fileNameTTNPart, i);
                    currentProgress += progressPart;




                    UI.progressBar.setValue(new Double(currentProgress).intValue());
                }
                UI.progressBar.setValue(100);
                UI.setStatusText("Статус - генерация завершена");
            } catch (Exception e) {}
            return null;
        }

        private void fillLogFile(String fileName, int index) {
            for (String error : errors.keySet()) {
                String s = "Строка данных = " + (index+1) + ", Файл - " + fileName + ".xls" + ", " + error + " : " + errors.get(error);
                try {
                    BufferedWriter bw = new BufferedWriter(new FileWriter(logFile, true));
                    bw.append('\n');
                    bw.append(s);
                    bw.close();
                } catch (IOException e) {
                }
            }
            errors.clear();
        }

        private void createLogFile(String directoryPath) {
            logFile = new File(directoryPath + File.separatorChar + "log.txt");
            String s = "Log";
            try {
                BufferedWriter bw = new BufferedWriter(new FileWriter(logFile, true));
                bw.append(s);
                bw.close();
            } catch (IOException e) {
            }
        }

        private String getFileNameTTNPart(int i, Cell ttnNumber) {
            String fileNameTTNPart = "";
            if (ttnNumber.getCellType() == CellType.NUMERIC) {
                fileNameTTNPart = String.valueOf((int) ttnNumber.getNumericCellValue());
            } else if (ttnNumber.getCellType() == CellType.STRING) {
                fileNameTTNPart = ttnNumber.getRichStringCellValue().getString();
            }
            if (!isValidName(fileNameTTNPart)) {
                errors.put("Ошибка названия файла", "недопустимые симовлы " + fileNameTTNPart + " назван = TTN-" + (i + 1));
                fileNameTTNPart = "TTN-" + (i + 1);
            }
            return fileNameTTNPart;
        }

        private static void cellCopy(Cell cell, Cell resCell, CellMapping mapping) {
            if (cell == null) {
                resCell.setBlank();
                return;
            }

            if (mapping.getReplaceStuff()) {
                if (cell.getCellType() == CellType.STRING) {
                    String replaceResult = CellMapping.replaceStuff(mapping, cell.getRichStringCellValue().getString());
                    if (replaceResult.equals("UNKNOWN_RESULT")) {
                        errors.put("Ошибка замены", "нет данных для замены");
                        resCell.setCellValue("ОШИБКА");
                    } else {
//                        String cellId = mapping.getColumn() + (mapping.getRow() + 1);
//                        if (cellId.equals("N4")) {
//                            Font font = resCell.getSheet().getWorkbook().createFont();
//                            font.setFontHeight((short)100);
//                            resCell.getCellStyle().setFont(font);
//                            resCell.getCellStyle().setWrapText(true);
//                        }

//                        resCell.getCellStyle().setShrinkToFit(true);
//                        resCell.getRow().setHeight((short) -1);

                        resCell.setCellValue(replaceResult);
                    }
                } else {
                    errors.put("Ошибка замены", "нет данных для замены");
                    resCell.setCellValue("ОШИБКА");
                }
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
                case ERROR:
                    errors.put("Ошибка клетки", "плохие данные в столбце "+ mapping.getDataColumn());
                    resCell.setCellValue("ОШИБКА");
                    break;
            }
        }

        private static void fillSheet(Integer sheetNumber, Sheet sheetToFill, Row dataRow) {
            for (CellMapping cellMapping : CellMapping.mapping.get(sheetNumber)) {
                Row resultRow = sheetToFill.getRow(cellMapping.getRow());
                Cell resultCell = resultRow.getCell(CellReference.convertColStringToIndex(cellMapping.getColumn()));
                cellCopy(dataRow.getCell(CellReference.convertColStringToIndex(cellMapping.getDataColumn())), resultCell, cellMapping);
            }
        }
    }
}
