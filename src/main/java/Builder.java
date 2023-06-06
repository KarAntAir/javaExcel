import javax.swing.*;
import java.awt.*;
import java.awt.event.ActionListener;
import java.awt.event.WindowAdapter;
import java.awt.event.WindowEvent;

public class Builder {
    private JFrame mainFrame;
    private JPanel dataSelectorPanel;
    private JPanel dataSelectorPathPanel;
    private JPanel statusBarPanel;
    private JPanel actionPanel;
    public static JProgressBar progressBar;
    private JLabel statusText;
    public void prepareGUI(){
        mainFrame = new JFrame("UI");
        mainFrame.setSize(500,500);
        mainFrame.setLayout(new GridLayout(4, 1));

        mainFrame.addWindowListener(new WindowAdapter() {
            public void windowClosing(WindowEvent windowEvent){
                System.exit(0);
            }
        });
        dataSelectorPanel = new JPanel();
        dataSelectorPathPanel = new JPanel();
        statusBarPanel = new JPanel();
        actionPanel = new JPanel();

        mainFrame.add(dataSelectorPanel);
        mainFrame.add(dataSelectorPathPanel);
        mainFrame.add(statusBarPanel);
        mainFrame.add(actionPanel);
        mainFrame.setVisible(true);
    }

    public void showSwingUI(){
        JPanel dataSelectorPan = new JPanel();
        JLabel dataSelectorText = new JLabel("Выберете файл данных");
        JButton dataSelectorButton = new JButton("Выбрать файл данных");
        dataSelectorPan.add(dataSelectorText);
        dataSelectorPan.add(dataSelectorButton);

        JPanel dataPathPan = new JPanel();
        JLabel dataResultText = new JLabel("Файл не выбран");
        dataPathPan.add(dataResultText);
        dataSelectorPathPanel.add(dataPathPan);

        ActionListener templateL = new DataListener(dataResultText);
        dataSelectorButton.addActionListener(templateL);
        dataSelectorPanel.add(dataSelectorPan);

        JPanel barPan = new JPanel();
        progressBar = new JProgressBar();
        progressBar.setValue(0);
        progressBar.setStringPainted(true);
        barPan.add(progressBar);
        statusBarPanel.add(barPan);

        JPanel actionPan = new JPanel();
        JButton actionButton = new JButton("Run");
        actionPan.add(actionButton);
        statusText = new JLabel("Статус - генерация не начата");
        actionPan.add(statusText);
        ActionListener runListener = new RunListener();
        actionButton.addActionListener(runListener);
        actionPanel.add(actionPan);

        mainFrame.setVisible(true);
    }

    public void setStatusText(String text) {
        statusText.setText(text);
    }
}
