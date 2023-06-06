import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;

public class RunListener implements ActionListener {
    @Override
    public void actionPerformed(ActionEvent e) {
        try {
            JavaExcel.UI.setStatusText("Статус - генерация в процессе");
            ExcelCellNumbers.fillRowsDataList();
            JavaExcel.doWork();
        } catch (Exception ex) {
            throw new RuntimeException(ex);
        }
    }
}