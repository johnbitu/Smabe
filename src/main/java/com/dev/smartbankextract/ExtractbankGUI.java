package com.dev.smartbankextract;

import javax.swing.*;
import javax.swing.filechooser.FileNameExtensionFilter;
import java.awt.*;
import java.io.File;

import java.util.logging.Logger;

public class ExtractbankGUI {
    private static final Logger logger = Logger.getLogger(ExtractbankGUI.class.getName());

    public static void startGUI() {
        JFrame frame = new JFrame("SmartBank Extract");
        frame.setSize(600, 300);
        frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        frame.setLayout(new GridLayout(4, 2, 10, 10));

        JTextField filePathField = new JTextField();
        filePathField.setEditable(false);

        JButton browseButton = new JButton("Selecionar Planilha");
        browseButton.addActionListener(e -> {
            JFileChooser fileChooser = new JFileChooser();
            fileChooser.setFileFilter(new FileNameExtensionFilter("Planilhas XLSX", "xlsx"));
            fileChooser.setCurrentDirectory(new File("Extratos/"));

            int returnValue = fileChooser.showOpenDialog(null);
            if (returnValue == JFileChooser.APPROVE_OPTION) {
                filePathField.setText(fileChooser.getSelectedFile().getAbsolutePath());
                logger.info("Arquivo selecionado: " + fileChooser.getSelectedFile().getAbsolutePath());
            }
        });

        JTextField spreadsheetIdField = new JTextField();
        JButton processButton = new JButton("Processar");
        JTextArea logArea = new JTextArea();
        logArea.setEditable(false);
        JScrollPane scrollPane = new JScrollPane(logArea);

        processButton.addActionListener(e -> {
            String filePath = filePathField.getText();
            String spreadsheetId = spreadsheetIdField.getText();

            try {
                logger.info("Processando arquivo: " + filePath + " com ID da planilha: " + spreadsheetId);
                SmartbankextractApplication.readPlanilha(filePath, spreadsheetId);
                logArea.append("Processamento concluído com sucesso!\n");
            } catch (Exception ex) {
                logger.severe("Erro ao processar: " + ex.getMessage());
                logArea.append("Erro: " + ex.getMessage() + "\n");
            }
        });

        frame.add(new JLabel("Caminho do Arquivo:"));
        frame.add(filePathField);
        frame.add(browseButton);
        frame.add(new JLabel("ID da Planilha do Google:"));
        frame.add(spreadsheetIdField);
        frame.add(new JLabel());
        frame.add(processButton);
        frame.add(scrollPane);

        frame.setVisible(true);
    }
}
