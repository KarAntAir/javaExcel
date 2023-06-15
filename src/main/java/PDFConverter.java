import java.io.BufferedReader;
import java.io.File;
import java.io.IOException;
import java.io.InputStreamReader;

public class PDFConverter {

    public static void saveAsPdf2() throws InterruptedException, IOException {

// write the workbook as a temporary file
// (if you don't start from a workbook, this step might differ)
        File tempExcelFile = new File("5918.xlsx");

// call libreoffice headless and politely
// ask it to convert our xlsm file to pdf
        ProcessBuilder pb = new ProcessBuilder(
                "/Applications/LibreOffice.app/Contents/MacOS/soffice", "--headless",
                "--convert-to", "pdf", tempExcelFile.getAbsolutePath(),
                "--outdir", "/Users/ak/workspace/Personal_Projects/javaExcel2/javaExcel"
        );
//        ProcessBuilder pb = new ProcessBuilder(
//                "libreoffice", "--headless",
//                "--convert-to", "pdf", tempExcelFile.getAbsolutePath(),
//                "--outdir", tempDir.toAbsolutePath().toString()
//        );
        pb.redirectErrorStream(true);
        Process process = pb.start();
        BufferedReader reader = new BufferedReader(
                new InputStreamReader(process.getInputStream())
        );
        String line;
        while ((line = reader.readLine()) != null) {
            System.out.println("[libreoffice stdout+stderr] " + line);
        }
        process.waitFor();
        System.out.println("converted");

// now the file has been converted

// read the converted file and send/use it

// remove the temp dir
    }

}