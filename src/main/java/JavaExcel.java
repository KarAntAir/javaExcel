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
        new JavaExcel();
//        String s1 = "ООО \"КОРОНА РУС\" Москва, ул. Каспийская, д. 22, корпус 1, стр. 5, этаж 5, пом. 9, к. 17А, офис 86 телефон: +7\u00A0(925)\u00A0641-85-57 р/с 40702810801840000325 АО \"АЛЬФА-БАНК\" к/с 30101810200000000593, БИК 044525593";
//        String s2 = "ООО \"КОРОНА РУС\" Москва, ул. Каспийская, д. 22, корпус 1, стр. 5, этаж 5, пом. 9, к. 17А, офис 86 телефон: +7 (925) 641-85-57 р/с 40702810801840000325 АО \"АЛЬФА-БАНК\" к/с 30101810200000000593, БИК 044525593";
//        System.out.println(s1.length());
//        System.out.println(s2.length());
//        String s = "ООО \"КОРОНА РУС\" Москва, ул. Каспийская, д. 22, корпус 1, стр. 5, этаж 5, пом. 9, к. 17А, офис 86 телефон: +7 (925) 641-85-57 р/с 40702810801840000325 АО \"АЛЬФА-БАНК\" к/с 30101810200000000593, БИК 044525593ООО \"КОРОНА РУС\" Москва, ул. Каспийская, д. 22, корпус 1, стр. 5, этаж 5, пом. 9, к. 17А, офис 86 телефон: +7 (925) 641-85-57 р/с 40702810801840000325 АО \"АЛЬФА-БАНК\" к/с 30101810200000000593, БИК 044525593";
//        new JavaExcel();
//        System.out.println(s);
//        System.out.println(s1);


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
}
