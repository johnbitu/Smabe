package com.dev.smartbankextract;

import javax.swing.*;
import javax.swing.filechooser.FileNameExtensionFilter;
import java.awt.*;
import java.io.File;
import java.util.logging.Logger;

public class ExtractbankGUI {
    private static final Logger logger = Logger.getLogger(ExtractbankGUI.class.getName());

    public static void startGUI() {
        // Configuração inicial do frame
        JFrame frame = new JFrame("Smabe");
        frame.setSize(600, 400);
        frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        frame.setLayout(new GridBagLayout());
        frame.getContentPane().setBackground(new Color(40, 44, 52)); // Fundo escuro

        // Configuração da fonte
        Font customFont = new Font("JetBrains Mono", Font.PLAIN, 14);

        // Configurações de layout
        GridBagConstraints gbc = new GridBagConstraints();
        gbc.fill = GridBagConstraints.HORIZONTAL;
        gbc.insets = new Insets(10, 10, 10, 10);

        // Campo para exibir o caminho do arquivo
        JLabel filePathLabel = new JLabel("Caminho do Arquivo:");
        filePathLabel.setForeground(Color.WHITE);
        filePathLabel.setFont(customFont);
        gbc.gridx = 0;
        gbc.gridy = 0;
        frame.add(filePathLabel, gbc);

        JTextField filePathField = new JTextField();
        filePathField.setEditable(false);
        filePathField.setFont(customFont);
        filePathField.setBackground(new Color(30, 34, 40));
        filePathField.setForeground(Color.WHITE);
        gbc.gridx = 1;
        gbc.gridy = 0;
        gbc.gridwidth = 2;
        frame.add(filePathField, gbc);

        // Botão para selecionar a planilha
        JButton browseButton = new JButton("Selecionar Planilha");
        browseButton.setFont(customFont);
        browseButton.setBackground(new Color(58, 63, 73));
        browseButton.setForeground(Color.WHITE);
        browseButton.setFocusPainted(false);
        browseButton.addActionListener(e -> {
            JFileChooser fileChooser = new JFileChooser();
            fileChooser.setFileFilter(new FileNameExtensionFilter("Somente XLSX, XLS e CSV", "xlsx","csv","xls"));
            fileChooser.setCurrentDirectory(new File("Extratos/"));

            int returnValue = fileChooser.showOpenDialog(null);
            if (returnValue == JFileChooser.APPROVE_OPTION) {
                filePathField.setText(fileChooser.getSelectedFile().getAbsolutePath());
                logger.info("Arquivo selecionado: " + fileChooser.getSelectedFile().getAbsolutePath());
            }
        });
        gbc.gridx = 3;
        gbc.gridy = 0;
        gbc.gridwidth = 1;
        frame.add(browseButton, gbc);

        // Campo para inserir o ID da planilha do Google
        JLabel spreadsheetIdLabel = new JLabel("ID da Planilha do Google:");
        spreadsheetIdLabel.setForeground(Color.WHITE);
        spreadsheetIdLabel.setFont(customFont);
        gbc.gridx = 0;
        gbc.gridy = 1;
        frame.add(spreadsheetIdLabel, gbc);

        JTextField spreadsheetIdField = new JTextField();
        spreadsheetIdField.setFont(customFont);
        spreadsheetIdField.setBackground(new Color(30, 34, 40));
        spreadsheetIdField.setForeground(Color.WHITE);
        gbc.gridx = 1;
        gbc.gridy = 1;
        gbc.gridwidth = 3;
        frame.add(spreadsheetIdField, gbc);


        // Área de logs
        JTextArea logArea = new JTextArea();
        logArea.setFont(customFont);
        logArea.setBackground(new Color(30, 34, 40));
        logArea.setForeground(Color.WHITE);
        logArea.setEditable(false);
        JScrollPane scrollPane = new JScrollPane(logArea);
        scrollPane.setBorder(BorderFactory.createLineBorder(new Color(58, 63, 73)));
        gbc.gridx = 0;
        gbc.gridy = 3;
        gbc.gridwidth = 4;
        gbc.weightx = 1;
        gbc.weighty = 1;
        gbc.fill = GridBagConstraints.BOTH;
        frame.add(scrollPane, gbc);


        // Botão para processar
        JButton processButton = new JButton("Processar");
        processButton.setFont(customFont);
        processButton.setBackground(new Color(58, 63, 73));
        processButton.setForeground(Color.WHITE);
        processButton.setFocusPainted(false);
        processButton.addActionListener(e -> {
            String filePath = filePathField.getText();
            String spreadsheetId = spreadsheetIdField.getText();

//            JTextArea logArea = new JTextArea();
//            logArea.setEditable(false);

            try {
                logger.info("Processando arquivo: " + filePath + " com ID da planilha: " + spreadsheetId);
                if (filePath.endsWith(".csv")) {
                    SmartbankextractApplication.readCsv(filePath, spreadsheetId);
                } else {
                    SmartbankextractApplication.readPlanilha(filePath, spreadsheetId);
                }
                logArea.append("Processamento concluído com sucesso!\n");
            } catch (Exception ex) {
                logger.severe("Erro ao processar: " + ex.getMessage());
                logArea.append("Erro: " + ex.getMessage() + "\n");
            }
        });
        gbc.gridx = 1;
        gbc.gridy = 2;
        gbc.gridwidth = 2;
        gbc.gridheight = 1;
        frame.add(processButton, gbc);



        // Exibir a janela
        frame.setVisible(true);
    }
}
