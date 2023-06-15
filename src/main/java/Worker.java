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
    private static HashMap<String, String> errors = new HashMap<String, String>();
    private static File logFile;
    private static Date currentWorkbookDate;

    private static RuleBasedNumberFormat nf = new RuleBasedNumberFormat(Locale.forLanguageTag("ru"), RuleBasedNumberFormat.SPELLOUT);
    private static File TTNFile = new File("ttn.xls");

    private Boolean isValidName(String text) {
        final Character[] INVALID_WINDOWS_SPECIFIC_CHARS = {':', '\\', '/', '"', '*', '<', '>', '?', '|'};
        return Arrays.stream(INVALID_WINDOWS_SPECIFIC_CHARS)
                .noneMatch(ch -> text.contains(ch.toString()));
    }

    @Override
    protected Object doInBackground() {
        int startIndex = 0;
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

                File fileTTNRes = new File(newDirectory + File.separatorChar + fileNameTTNPart + ".xls");
                Files.copy(TTNFile.toPath(), fileTTNRes.toPath(), StandardCopyOption.REPLACE_EXISTING);           //копируем шаблон для дальнейшего заполнения

                FileInputStream inputStream = new FileInputStream(fileTTNRes);

                Workbook workbook = new HSSFWorkbook(inputStream);

                Sheet sheet1 = workbook.getSheetAt(0); // заполняем первый лист накладной

                //date stuff
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

                inputStream.close();

                FileOutputStream os = new FileOutputStream(fileTTNRes);
                workbook.write(os);
                workbook.close();
                os.close();
                fillLogFile(fileNameTTNPart, i);
                currentProgress += progressPart;

                currentWorkbookDate = null;
                Builder.progressBar.setValue(new Double(currentProgress).intValue());
            }
            Builder.progressBar.setValue(100);
            JavaExcel.UI.setStatusText("Статус - генерация завершена");
        } catch (Exception e) {
            System.out.println("main ex");
            System.out.println(e.getMessage());
            e.printStackTrace();
        }
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
            System.out.println("cell "+ mapping.getColumn() + mapping.getOriginalRow());
            System.out.println(e.getMessage());
        }
    }

    private static void fillSheet(Integer sheetNumber, Sheet sheetToFill, Row dataRow) {
        for (CellMapping cellMapping : CellMapping.mapping.get(sheetNumber)) {
            Row resultRow = sheetToFill.getRow(cellMapping.getIndexRow());
            Cell resultCell = resultRow.getCell(CellReference.convertColStringToIndex(cellMapping.getColumn()));
            cellCopy(dataRow.getCell(CellReference.convertColStringToIndex(cellMapping.getDataColumn())), resultCell, cellMapping);
        }
    }

    private static String getMonthInWords(String month) {
        switch (month) {
            case "01": return "января";
            case "02": return "февраля";
            case "03": return "марта";
            case "04": return "апреля";
            case "05": return "мая";
            case "06": return "июня";
            case "07": return "июля";
            case "08": return "августа";
            case "09": return "сентября";
            case "10": return "октября";
            case "11": return "ноября";
            case "12": return "декабря";
        }
        errors.put("ошибка парсинга месяца", month);
        return "ошибка";
    }
}