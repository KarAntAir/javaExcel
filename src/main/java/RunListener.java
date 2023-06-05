import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;

public class RunListener implements ActionListener {
    @Override
    public void actionPerformed(ActionEvent e) {
        try {
            System.out.println("run list click");
            ExcelCellNumbers.fillRowsDataList();
            System.out.println("second");

            JavaExcel.doWork();
        } catch (Exception ex) {
            System.out.println(ex.getMessage());
            throw new RuntimeException(ex);
        }
    }
}