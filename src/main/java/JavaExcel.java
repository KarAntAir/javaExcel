import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellReference;

import javax.swing.*;
import java.awt.*;
import java.io.*;
import java.nio.file.Files;
import java.nio.file.StandardCopyOption;
import java.text.SimpleDateFormat;
import java.util.*;

public class JavaExcel {
    public static File TTNFile = new File("ttn.xls");
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
        new JavaExcel();
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
            double endIndex = ExcelCellNumbers.rowsDataList.size();
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

                    File fileTTNRes = new File(newDirectory + File.separatorChar + (i + 1) + "tempFile.xls");
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

                    inputStream.close();

                    fileTTNRes = new File(newDirectory + File.separatorChar + (i + 1) + "tempFile.xls");
                    File renamedFile = new File(newDirectory + File.separatorChar + fileNameTTNPart + ".xls");
                    fileTTNRes.renameTo(renamedFile);

                    FileOutputStream os = new FileOutputStream(renamedFile);
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
