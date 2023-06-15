import com.documents4j.api.DocumentType;
import com.documents4j.api.IConverter;
import com.documents4j.job.LocalConverter;

import java.io.*;
import java.util.concurrent.TimeUnit;


public class PDFConverter {
    public static void saveAsPDF() throws IOException {
        IConverter converter = LocalConverter.builder()
                .baseFolder(new File("/Users/ak/workspace/Personal_Projects/javaExcel2/javaExcel"))
                .workerPool(20, 25, 2, TimeUnit.SECONDS)
                .processTimeout(5, TimeUnit.SECONDS)
                .build();
        converter.convert(new File("5918.xls"))
                .as(DocumentType.PDF)
                .to(new File("output.pdf"));

    }

    public static void saveAsPdf2() throws IOException, InterruptedException {

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
