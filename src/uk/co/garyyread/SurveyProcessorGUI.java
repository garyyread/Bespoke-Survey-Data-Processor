package uk.co.garyyread;

import java.awt.Dimension;
import java.awt.GridBagConstraints;
import java.awt.GridBagLayout;
import java.awt.event.ActionEvent;
import java.awt.event.WindowAdapter;
import java.awt.event.WindowEvent;
import java.awt.event.WindowListener;
import java.io.IOException;
import java.util.ArrayList;
import java.util.logging.Level;
import java.util.logging.Logger;
import javax.swing.BoxLayout;
import javax.swing.Icon;
import javax.swing.JButton;
import javax.swing.JCheckBox;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JPanel;
import javax.swing.JScrollPane;
import javax.swing.JTabbedPane;
import javax.swing.JTextField;
import javax.swing.SwingUtilities;
import jxl.write.WriteException;

/**
 * Graphical interface for use with the SurveyProcessor class, a bespoke java
 * application developed for a university student at Swansea University in aid
 * of their masters project.
 *
 * @author Gary
 * @since 2015
 */
public class SurveyProcessorGUI extends JFrame {

    private final SurveyProcessor sp;

    private JLabel inputLabel;
    private JTextField inputField;
    private JButton inputButton;
    private JLabel outputLabel;
    private JTextField outputField;
    private JButton outputButton;
    private ArrayList<JCheckBox> sheetCheckBoxArray;
    private JCheckBox debugCheckBox;
    private JButton processButton;
    private JTabbedPane tabbedPane;
    private JPanel debugTab;
    private JPanel processTab;
    private JPanel sheetCheckBoxPanel;

    private final String inputLabelStr = "Source File";
    private final String outputLabelStr = "Output File";
    private final String debugCheckBoxStr = "Debug";
    private final String processButtonStr = "Process";
    private final String debugTabStr = "Debug Dialog";
    private final String processTabStr = "Process Dialog";
    private final Dimension WINDOW_SIZE = new Dimension(400, 400);
    private final String inputButtonStr = "Load";
    private final String outputButtonStr = "Create";

    public SurveyProcessorGUI() {
        sp = new SurveyProcessor();

        initGUI();
        showGUI();
    }

    private void initGUI() {
        inputLabel = new JLabel();
        inputLabel.setText(inputLabelStr);
        inputField = new JTextField();
        inputButton = new JButton(inputButtonStr);
        inputButton.addActionListener((ActionEvent e)-> {
            inputButtonAction();
        });

        outputLabel = new JLabel();
        outputLabel.setText(outputLabelStr);
        outputField = new JTextField();
        outputButton = new JButton(outputButtonStr);
        outputButton.addActionListener((ActionEvent e)-> {
            outputButtonAction();
        });

        debugCheckBox = new JCheckBox();
        debugCheckBox.setText(debugCheckBoxStr);
        debugCheckBox.setSelected(false);
        debugCheckBox.addActionListener((ActionEvent e) -> {
            sp.setDebugging(debugCheckBox.isSelected());
        });

        processButton = new JButton();
        processButton.setText(processButtonStr);
        processButton.addActionListener((ActionEvent e) -> {
            boolean display = false;
            for (JCheckBox box : sheetCheckBoxArray) {
                if (box.isSelected()) {
                    try {
                        sp.processSheet(box.getText());
                        display = true;
                    } catch (WriteException | IOException ex) {
                        Logger.getLogger(SurveyProcessorGUI.class.getName()).log(Level.SEVERE, null, ex);
                    }
                }
            }
            if (display) {
                sp.displayMessage("Finished!");
            }
        });
        
        processTab = new JPanel();
        processTab.setLayout(new GridBagLayout());
        GridBagConstraints c = new GridBagConstraints();
        
        sheetCheckBoxPanel = new JPanel();
        sheetCheckBoxPanel.setLayout(new BoxLayout(sheetCheckBoxPanel, BoxLayout.Y_AXIS));
        JScrollPane scrollPane = new JScrollPane(sheetCheckBoxPanel);
        
        c.gridx = 0;
        c.gridy = 0;
        c.weightx = 0;
        c.weighty = 0;
        c.fill = GridBagConstraints.NONE;
        processTab.add(inputLabel, c);
        c.gridx = 1;
        c.gridy = 0;
        c.weightx = 2;
        c.weighty = 0;
        c.fill = GridBagConstraints.HORIZONTAL;
        processTab.add(inputField, c);
        c.gridx = 2;
        c.gridy = 0;
        c.weightx = 0;
        c.weighty = 0;
        c.fill = GridBagConstraints.NONE;
        processTab.add(inputButton, c);
        
        c.gridx = 0;
        c.gridy = 1;
        c.weightx = 0;
        c.weighty = 0;
        c.fill = GridBagConstraints.NONE;
        //processTab.add(outputLabel, c);
        c.gridx = 1;
        c.gridy = 1;
        c.weightx = 2;
        c.fill = GridBagConstraints.HORIZONTAL;
        //processTab.add(outputField, c);
        c.gridx = 2;
        c.gridy = 1;
        c.weightx = 1;
        c.weighty = 0;
        c.fill = GridBagConstraints.HORIZONTAL;
        //processTab.add(outputButton, c);
        
        c.gridx = 0;
        c.gridy = 2;
        c.weightx = 1;
        c.weighty = 1;
        c.fill = GridBagConstraints.BOTH;
        processTab.add(scrollPane, c);
        
        c.gridx = 1;
        c.gridy = 3;
        c.weightx = 1;
        c.weighty = 0;
        c.fill = GridBagConstraints.HORIZONTAL;
        processTab.add(processButton, c);
        c.gridx = 0;
        c.gridy = 3;
        c.weightx = 0;
        c.weighty = 0;
        c.fill = GridBagConstraints.NONE;
        processTab.add(debugCheckBox, c);
        
        
        add(processTab);
        setPreferredSize(WINDOW_SIZE);
        setDefaultCloseOperation(JFrame.DISPOSE_ON_CLOSE);
        setTitle("Survey Processor - G.Read");
        addWindowListener(new WindowAdapter() {
            @Override
            public void windowClosing(WindowEvent e) {
                //NA
            }
        });
    }

    private void showGUI() {
        SwingUtilities.invokeLater(() -> {
            pack();
            setVisible(true);
        });
    }

    public static void main() {
        SurveyProcessorGUI gui = new SurveyProcessorGUI();
    }

    private void inputButtonAction() {
        if (sp.loadWorkbook(inputField.getText())) {
        
            sheetCheckBoxPanel.removeAll();
            sheetCheckBoxArray = new ArrayList<>();
            for (String sheetName : sp.getSheetNames()) {
                JCheckBox cb = new JCheckBox(sheetName, false);
                sheetCheckBoxArray.add(cb);
                sheetCheckBoxPanel.add(cb);
            }
        }
        sheetCheckBoxPanel.revalidate();
    }

    private void outputButtonAction() {
        try {
            sp.createWritableWorkbook(outputField.getText());
        } catch (IOException ex) {
            Logger.getLogger(SurveyProcessorGUI.class.getName()).log(Level.SEVERE, null, ex);
        }
    }
}
