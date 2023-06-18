import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;

public class RunListener implements ActionListener {
    @Override
    public void actionPerformed(ActionEvent e) {
        try {
            ExcelCellNumbers.fillRowsDataList();
            JavaExcel.doWork();
        } catch (Exception ex) {
            throw new RuntimeException(ex);
        }
    }
}