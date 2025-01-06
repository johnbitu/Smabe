package com.dev.smartbankextract;

import javax.swing.*;
import javax.swing.filechooser.FileNameExtensionFilter;
import java.awt.*;
import java.io.File;

public class ExtractbankGUI {

    private JFileChooser fileChooser;
    private static FileNameExtensionFilter filter = new FileNameExtensionFilter("Selecione apenas XLSX e XLS", "xlsx", "xls");



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
            fileChooser = new JFileChooser();
            fileChooser.setFileFilter(filter); // Configura o filtro para arquivos XLSX e XLS
            fileChooser.setCurrentDirectory(new File("Extratos/"));

            int returnValue = fileChooser.showOpenDialog(null);
            if (returnValue == JFileChooser.APPROVE_OPTION) {
                filePathField.setText(fileChooser.getSelectedFile().getAbsolutePath());
            }
        });

        JTextField spreadsheetIdField = new JTextField();

        JButton processButton = new JButton("Processar");
        JTextArea logArea = new JTextArea();
        logArea.setEditable(false); // Impedir edição manual nos logs
        JScrollPane scrollPane = new JScrollPane(logArea);

        processButton.addActionListener(e -> {
            String filePath = filePathField.getText();
            String spreadsheetId = spreadsheetIdField.getText();

            try {
                logArea.append("Caminho do Arquivo: " + filePath + "\n");
                logArea.append("ID da Planilha: " + spreadsheetId + "\n");
                logArea.append("Processamento concluído com sucesso!\n");
            } catch (Exception ex) {
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