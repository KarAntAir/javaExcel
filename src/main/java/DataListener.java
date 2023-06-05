import javax.swing.*;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;

public class DataListener implements ActionListener {

    private JLabel label;
    public DataListener(JLabel l) {
        this.label = l;
    }
    @Override
    public void actionPerformed(ActionEvent e) {
        File workingDirectory = new File(System.getProperty("user.dir"));
        JFileChooser j = new JFileChooser();
        j.setCurrentDirectory(workingDirectory);
        int result = j.showOpenDialog(null);
        if (result == JFileChooser.APPROVE_OPTION) {
            label.setText(j.getSelectedFile().getAbsolutePath());
            JavaExcel.dataFile = new File(j.getSelectedFile().getAbsolutePath());
        } else {
            label.setText("Отменено");
        }
    }
}
