import com.ibm.icu.text.RuleBasedNumberFormat;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellReference;

import javax.swing.*;
import java.io.*;
import java.math.BigDecimal;
import java.math.RoundingMode;
import java.nio.file.Files;
import java.nio.file.StandardCopyOption;
import java.text.SimpleDateFormat;
import java.util.*;

public class Worker extends SwingWorker<Object, Object> {
    private static HashMap<String, String> errors = new HashMap<>();
    private static File logFile;
    private static Date currentWorkbookDate;
    private static HashMap<String, String> paths = new HashMap<>();
    private static RuleBasedNumberFormat nf = new RuleBasedNumberFormat(Locale.forLanguageTag("ru"), RuleBasedNumberFormat.SPELLOUT);
    private static File TTNFile = new File("ttn.xls");

    @Override
    protected Object doInBackground() {
        createDirectories();
        createLogFile(paths.get("base"));
        originalGeneration();
        double counter = updatedGeneration();
        PDFGeneration(counter);
        return null;
    }

    private void PDFGeneration(double counter) {
        JavaExcel.UI.setStatusText("Статус - генерация PDF в процессе");
        double progressPart = 100d / counter;
        double currentProgress = 0d;

        String sourceFiles = new File(paths.get("updated")).getAbsolutePath() + "/*.xls";
        String targetDir = new File(paths.get("pdfs")).getAbsolutePath() + "/";
        ProcessBuilder pb;
        try {
            String operationSystem = System.getProperty("os.name").toLowerCase();
            if (operationSystem.contains("win")) {
                pb = new ProcessBuilder(
                    "libreoffice", "--headless",
                    "--convert-to", "pdf", sourceFiles,
                    "--outdir", targetDir);
            } else {
                pb = new ProcessBuilder(
                        "/bin/bash", "-c",
                        "/Applications/LibreOffice.app/Contents/MacOS/soffice --headless --convert-to pdf "
                                + sourceFiles + " --outdir "
                                + targetDir
                );
            }
            Process process = pb.start();

            pb.redirectErrorStream(true);
            BufferedReader reader = new BufferedReader(
                    new InputStreamReader(process.getInputStream())
            );
            String line;
            while ((line = reader.readLine()) != null) {
                //System.out.println("[libreoffice stdout+stderr] " + line);
                currentProgress += progressPart;
                Builder.progressBar.setValue(new Double(currentProgress).intValue());
            }

            process.waitFor();
        } catch (Exception e) {
            errors.put("Ошибка генерации PDF", e.getMessage());
            fillLogFile();
            System.out.println(e.getMessage());
        }
        Builder.progressBar.setValue(100);
        JavaExcel.UI.setStatusText("Статус - генерация PDF завершена");

    }

    private double updatedGeneration() {
        JavaExcel.UI.setStatusText("Статус - генерация для PDF в процессе");

        File folder = new File(paths.get("original"));
        File[] listOfFiles = folder.listFiles();

        int startIndex = 0;
        double endIndex = listOfFiles.length;
        double progressPart = 100d / endIndex;
        double currentProgress = 0d;

        try {
            for (int i = startIndex; i < endIndex; i++) {

                if (listOfFiles[i].isFile()) {
                    File originalGenerated = listOfFiles[i];
                    File updatedGeneration = new File(paths.get("updated") + File.separatorChar + originalGenerated.getName());
                    Files.copy(originalGenerated.toPath(), updatedGeneration.toPath(), StandardCopyOption.REPLACE_EXISTING);

                    FileInputStream inputStream = new FileInputStream(updatedGeneration);

                    Workbook workbook = new HSSFWorkbook(inputStream);

                    Sheet sheet1 = workbook.getSheetAt(0);

                    resizeAllColumnsInSheet(sheet1, "GE");

                    Sheet sheet2 = workbook.getSheetAt(1);

                    resizeAllColumnsInSheet(sheet1, "GH");

                    applyPrintSettings(workbook, sheet1, sheet2);

                    inputStream.close();

                    try (FileOutputStream os = new FileOutputStream(updatedGeneration)){
                        workbook.write(os);
                        workbook.close();
                        os.flush();
                    }

                    currentProgress += progressPart;
                    Builder.progressBar.setValue(new Double(currentProgress).intValue());
                }
            }

            Builder.progressBar.setValue(100);
            JavaExcel.UI.setStatusText("Статус - генерация для PDF завершена");
        } catch (Exception e) {
            errors.put("Общая ошибка - 1", e.getMessage());
            fillLogFile();
        }
        return endIndex;
    }


    private void originalGeneration() {
        JavaExcel.UI.setStatusText("Статус - генерация в процессе");

        int startIndex = 0;
        double endIndex = ExcelCellNumbers.rowsDataList.size();
        double progressPart = 100d / endIndex;
        double currentProgress = 0d;

        try {
            for (int i = startIndex; i < endIndex; i++) {
                Row row = ExcelCellNumbers.rowsDataList.get(i);

                Cell ttnNumber = row.getCell(CellReference.convertColStringToIndex("L"));
                Cell nameCell = row.getCell(CellReference.convertColStringToIndex("Q"));
                String fileNameTTNPart = getFileNameTTNPart(i, ttnNumber, nameCell);

                File fileTTNRes = new File(paths.get("original") + File.separatorChar + fileNameTTNPart + ".xls");
                Files.copy(TTNFile.toPath(), fileTTNRes.toPath(), StandardCopyOption.REPLACE_EXISTING);           //копируем шаблон для дальнейшего заполнения

                FileInputStream inputStream = new FileInputStream(fileTTNRes);

                Workbook workbook = new HSSFWorkbook(inputStream);

                Sheet sheet1 = workbook.getSheetAt(0); // заполняем первый лист накладной

                Cell dateCell = row.getCell(12);
                if (dateCell.getCellType() == CellType.NUMERIC && DateUtil.isCellDateFormatted(dateCell)) {
                    currentWorkbookDate = dateCell.getDateCellValue();
                } else {
                    errors.put("Ошибка даты", "в данных не найдена дата");
                }

                fillSheet(1, sheet1, row);

                Sheet sheet2 = workbook.getSheetAt(1); // заполняем второй лист накладной

                fillSheet(2, sheet2, row);

                FormulaEvaluator formulaEvaluator = workbook.getCreationHelper().createFormulaEvaluator();
                formulaEvaluator.evaluateAll();

                applyPrintSettings(workbook, sheet1, sheet2);

                inputStream.close();

                try (FileOutputStream os = new FileOutputStream(fileTTNRes)){
                    workbook.write(os);
                    workbook.close();
                    os.flush();
                }

                fillLogFile(fileNameTTNPart, i);
                currentProgress += progressPart;

                currentWorkbookDate = null;
                Builder.progressBar.setValue(new Double(currentProgress).intValue());
            }
            Builder.progressBar.setValue(100);
            JavaExcel.UI.setStatusText("Статус - генерация завершена");
        } catch (Exception e) {
            errors.put("Общая ошибка - 2", e.getMessage());
            fillLogFile();
        }
    }

    private void createDirectories() {
        String timeStamp = new SimpleDateFormat("yyyyMMdd-HHmmss").format(Calendar.getInstance().getTime());
        paths.put("original", "result-"+timeStamp+"/originalGeneration");
        paths.put("updated", "result-"+timeStamp+"/updatedForPDFCreation");
        paths.put("pdfs", "result-"+timeStamp+"/pdfs");
        paths.put("base", "result-"+timeStamp);

        for (String newDir : paths.values()) {
            File dir = new File(newDir);
            if (!dir.exists()) {
                dir.mkdirs();
            }
        }
    }

    private static void fillLogFile(String fileName, int index) {
        for (String error : errors.keySet()) {
            String s = "Строка данных = " + (index + 1) + ", Файл - " + fileName + ".xls" + ", " + error + " : " + errors.get(error);
            writeLogLine(s);
        }
        errors.clear();
    }
    private static void fillLogFile() {
        for (String error : errors.keySet()) {
            String s = error + " : " + errors.get(error);
            writeLogLine(s);
        }
        errors.clear();
    }

    private static void writeLogLine(String s) {
        try {
            BufferedWriter bw = new BufferedWriter(new FileWriter(logFile, true));
            bw.append('\n');
            bw.append(s);
            bw.close();
        } catch (IOException e) {
        }
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

    private String getFileNameTTNPart(int i, Cell ttnNumber, Cell nameCell) {
        String fileNameTTNPart = "";
        if (nameCell.getCellType() == CellType.STRING) {
            fileNameTTNPart = nameCell.getRichStringCellValue().getString()
                    .replace(" ", "")
                    .replace(".", "") + "-";
        } else {
            fileNameTTNPart = "Фамилия-";
        }

        if (ttnNumber.getCellType() == CellType.NUMERIC) {
            fileNameTTNPart += String.valueOf((int) ttnNumber.getNumericCellValue());
        } else if (ttnNumber.getCellType() == CellType.STRING) {
            fileNameTTNPart += ttnNumber.getRichStringCellValue().getString();
        }
        if (!isValidName(fileNameTTNPart)) {
            errors.put("Ошибка названия файла", "недопустимые симовлы " + fileNameTTNPart + " назван = TTN-" + (i + 1));
            fileNameTTNPart = "TTN-" + (i + 1);
        }
        return fileNameTTNPart;
    }

    private static void cellCopy(Cell cell, Cell resCell, CellMapping mapping) {
        try {
            if (cell == null) {
                resCell.setBlank();
                return;
            }

            CellAction cellAction = mapping.getCellAction();

            if (cellAction == CellAction.REPLACE || cellAction == CellAction.REPLACE_RESIZE) {
                if (cell.getCellType() == CellType.STRING) {
                    String replaceResult = CellMapping.replaceStuff(mapping, cell.getRichStringCellValue().getString());
                    if (replaceResult.equals("UNKNOWN_RESULT")) {
                        errors.put("Ошибка замены", "нет данных для замены");
                        resCell.setCellValue("ОШИБКА");
                    } else {
                        resCell.setCellValue(replaceResult);
                        if (cellAction == CellAction.REPLACE_RESIZE) {
//                            resizeCellText(resCell);
                        }
                    }
                } else {
                    errors.put("Ошибка замены", "нет данных для замены");
                    resCell.setCellValue("ОШИБКА");
                }
            } else if (cellAction == CellAction.IN_WORDS) {
                String toWords = cell.getCellType() == CellType.NUMERIC ?
                        String.valueOf(cell.getNumericCellValue())
                        : cell.getRichStringCellValue().getString();
                BigDecimal bd = new BigDecimal(toWords).setScale(1, RoundingMode.DOWN);
                resCell.setCellValue(nf.format(bd));
            } else if ((cellAction == CellAction.DAY   // should be before other dates action
                    || cellAction == CellAction.MONTH
                    || cellAction == CellAction.YEAR)
                    && currentWorkbookDate == null) {
                resCell.setCellValue("ОШИБКА");
                errors.put("Ошибка даты", "в данных не найдена дата");
            } else if (cellAction == CellAction.DAY) {
                resCell.setCellValue(new SimpleDateFormat("dd").format(currentWorkbookDate));
            } else if (cellAction == CellAction.MONTH) {
                resCell.setCellValue(new SimpleDateFormat("MM").format(currentWorkbookDate));
            } else if (cellAction == CellAction.YEAR) {
                resCell.setCellValue(new SimpleDateFormat("yyyy").format(currentWorkbookDate));
            } else if (cellAction == CellAction.IN_WORDS_MONTH) {
                String m = new SimpleDateFormat("MM").format(currentWorkbookDate);
                resCell.setCellValue(getMonthInWords(m));
            } else if (cellAction == null) {
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
                        errors.put("Ошибка клетки", "плохие данные в столбце " + mapping.getDataColumn());
                        resCell.setCellValue("ОШИБКА");
                        break;
                }
            }
        } catch (Exception e) {
            errors.put("Общая ошибка - 3", e.getMessage());
            fillLogFile();
        }
    }

    private static void fillSheet(Integer sheetNumber, Sheet sheetToFill, Row dataRow) {
        for (CellMapping cellMapping : CellMapping.mapping.get(sheetNumber)) {
            Row resultRow = sheetToFill.getRow(cellMapping.getIndexRow());
            Cell resultCell = resultRow.getCell(CellReference.convertColStringToIndex(cellMapping.getColumn()));
            //TODO: resultCell may be null
            cellCopy(dataRow.getCell(CellReference.convertColStringToIndex(cellMapping.getDataColumn())), resultCell, cellMapping);
        }
    }

    private static String getMonthInWords(String month) {
        switch (month) {
            case "01":
                return "января";
            case "02":
                return "февраля";
            case "03":
                return "марта";
            case "04":
                return "апреля";
            case "05":
                return "мая";
            case "06":
                return "июня";
            case "07":
                return "июля";
            case "08":
                return "августа";
            case "09":
                return "сентября";
            case "10":
                return "октября";
            case "11":
                return "ноября";
            case "12":
                return "декабря";
        }
        errors.put("Ошибка парсинга месяца", month);
        return "ошибка";
    }

    private static void applyPrintSettings(Workbook workbook, Sheet sheet1, Sheet sheet2) {
        workbook.setPrintArea(0,  "$A$1:$GE$45");
        workbook.setPrintArea(1,  "$A$1:$GH$61");
        sheet1.setAutobreaks(true);
        sheet2.setAutobreaks(true);
    }
    private Boolean isValidName(String text) {
        final Character[] INVALID_WINDOWS_SPECIFIC_CHARS = {':', '\\', '/', '"', '*', '<', '>', '?', '|'};
        return Arrays.stream(INVALID_WINDOWS_SPECIFIC_CHARS)
                .noneMatch(ch -> text.contains(ch.toString()));
    }

    private static void resizeAllColumnsInSheet(Sheet sheet, String lastColumn) {
        for (int j = 0; j < CellReference.convertColStringToIndex(lastColumn); j++) {
            sheet.setColumnWidth(j, 222);
        }
    }
}