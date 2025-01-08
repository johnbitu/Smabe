package com.dev.smartbankextract;

import com.google.api.client.googleapis.auth.oauth2.GoogleCredential;
import com.google.api.services.sheets.v4.Sheets;
import com.google.api.services.sheets.v4.SheetsScopes;
import com.google.api.services.sheets.v4.model.ValueRange;
import io.github.cdimascio.dotenv.Dotenv;
import org.apache.poi.poifs.crypt.Decryptor;
import org.apache.poi.poifs.crypt.EncryptionInfo;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;

import java.io.BufferedReader;
import java.io.FileReader;
import java.util.ArrayList;
import java.util.List;

import java.awt.*;
import java.io.File;
import java.io.FileInputStream;
import java.math.BigDecimal;
import java.util.Arrays;
import java.util.Collections;

import java.util.logging.*;

@SpringBootApplication
public class SmartbankextractApplication {

	private static final Logger logger = Logger.getLogger(SmartbankextractApplication.class.getName());
	private static final Dotenv dotenv = Dotenv.configure().load();

	public static void main(String[] args) {
		configurarLogger();
		boolean isHeadless = GraphicsEnvironment.isHeadless();

		if (isHeadless) {
			logger.info("Executando em ambiente headless...");
			if (args.length < 2) {
				logger.severe("Parâmetros insuficientes. Uso: java -jar sua-aplicacao.jar <caminho-da-planilha> <id-da-planilha-google>");
				return;
			}

			String filePath = args[0];
			String spreadsheetId = args[1];

			try {
				if (filePath.endsWith(".csv")) {
					readCsv(filePath, spreadsheetId);
				} else if (filePath.endsWith(".xlsx")) {
					readPlanilha(filePath, spreadsheetId);
				} else {
					logger.severe("Formato de arquivo não suportado: " + filePath);
				}

				logger.info("Processamento concluído com sucesso!");
			} catch (Exception e) {
				logger.log(Level.SEVERE, "Erro durante o processamento", e);
			}
		} else {
			ExtractbankGUI.startGUI();
		}
	}

	public static void readCsv(String filePath, String spreadsheetId) throws Exception {
		logger.info("Lendo arquivo CSV: " + filePath);
		List<List<Object>> registros = new ArrayList<>();

		try (BufferedReader br = new BufferedReader(new FileReader(filePath))) {
			String line;
			while ((line = br.readLine()) != null) {
				String[] values = line.split(","); // Divide a linha usando vírgulas como delimitador

				// Adiciona os valores em uma lista de objetos
				List<Object> linha = new ArrayList<>();
				Collections.addAll(linha, values);
				registros.add(linha);
			}

			logger.info("CSV processado com sucesso. Enviando dados ao Google Sheets.");
			insertInGoogle(spreadsheetId, registros);
		} catch (Exception e) {
			logger.log(Level.SEVERE, "Erro ao processar o arquivo CSV", e);
			throw e;
		}
	}

	public static void readPlanilha(String filePath, String spreadsheetId) throws Exception {
		logger.info("Lendo planilha: " + filePath);
		String passwordPj = dotenv.get("PASSWORD_PJ");

		List<List<Object>> registros = new ArrayList<>();

		try (FileInputStream fileIs = new FileInputStream(new File(filePath));
			 POIFSFileSystem fileSystem = new POIFSFileSystem(fileIs)) {

			EncryptionInfo encryptionInfo = new EncryptionInfo(fileSystem);
			Decryptor decryptor = Decryptor.getInstance(encryptionInfo);

			if (!decryptor.verifyPassword(passwordPj)) {
				logger.severe("Senha incorreta para descriptografar a planilha.");
				throw new RuntimeException("Senha incorreta!");
			}

			try (Workbook workbook = new XSSFWorkbook(decryptor.getDataStream(fileSystem))) {
				Sheet sheet = workbook.getSheetAt(0);

				boolean dadosEncontrados = false;

				for (Row row : sheet) {
					// Identifica se a linha é onde começam os dados (ignorando cabeçalho)
					if (!dadosEncontrados && row.getCell(0) != null
							&& row.getCell(0).toString().equalsIgnoreCase("Data")) {
						dadosEncontrados = true;
						continue; // Ignora a linha do cabeçalho
					}

					if (dadosEncontrados) {
						// Lê apenas linhas que possuem dados nas colunas principais
						if (row.getCell(0) != null && !row.getCell(0).toString().isEmpty()) {
							List<Object> linha = new ArrayList<>();
							linha.add(row.getCell(0) != null ? row.getCell(0).toString() : ""); // Data
							linha.add(row.getCell(1) != null ? row.getCell(1).toString() : ""); // Título
							linha.add(row.getCell(2) != null ? row.getCell(2).toString() : ""); // Descrição
							String entrada = row.getCell(3) != null ? row.getCell(3).toString() : "";
							String saida = row.getCell(4) != null ? row.getCell(4).toString() : "";

							BigDecimal valorEntrada = entrada.isEmpty() ? BigDecimal.ZERO : new BigDecimal(entrada);
							BigDecimal valorSaida = saida.isEmpty() ? BigDecimal.ZERO : new BigDecimal(saida).negate();

							linha.add(valorEntrada); // Entrada
							linha.add(valorSaida);

							registros.add(linha);
						}
					}
				}

				logger.info("Planilha processada com sucesso. Enviando dados ao Google Sheets.");
				insertInGoogle(spreadsheetId, registros);
			}
		} catch (Exception e) {
			logger.log(Level.SEVERE, "Erro ao processar a planilha", e);
			throw e;
		}
	}


	public static void insertInGoogle(String spreadsheetId, List<List<Object>> registros) throws Exception {
		logger.info("Iniciando envio ao Google Sheets para ID: " + spreadsheetId);
		String range = "Dados1!A1"; // Define o intervalo inicial
		Sheets sheetsService = getSheetsService();

		// Monta o layout da planilha
		List<List<Object>> planilhaComLayout = new ArrayList<>();

		// Adiciona o cabeçalho
		List<Object> cabecalho = Arrays.asList("Data", "Titulo", "Descrição", "Entrada", "Saída", "Categoria", "Observações");
		planilhaComLayout.add(cabecalho);

		// Adiciona os dados lidos no layout
		for (List<Object> registro : registros) {
			List<Object> linhaFormatada = new ArrayList<>();

			// Supondo que os dados originais são organizados como:
			// Data, Descrição, Entrada, Saída
			linhaFormatada.add(registro.size() > 0 ? registro.get(0) : ""); // Data
			linhaFormatada.add(registro.size() > 1 ? registro.get(1) : ""); // Titulo
			linhaFormatada.add(registro.size() > 2 ? registro.get(2) : ""); // Descrição

			// Organiza os valores de Entrada e Saída
			String entrada = registro.size() > 3 ? registro.get(3).toString() : "";
			String saida = registro.size() > 4 ? registro.get(4).toString() : "";
			linhaFormatada.add(!entrada.isEmpty() ? entrada : ""); // Entrada
			linhaFormatada.add(!saida.isEmpty() ? saida : ""); // Saída

			// Adiciona campos de Categoria e Observações como vazios
			linhaFormatada.add(""); // Categoria
			linhaFormatada.add(""); // Observações

			planilhaComLayout.add(linhaFormatada);
		}

		// Envia os dados ao Google Sheets
		try {
			ValueRange body = new ValueRange().setValues(planilhaComLayout);

			sheetsService.spreadsheets().values()
					.update(spreadsheetId, range, body)
					.setValueInputOption("RAW")
					.execute();

			logger.info("Dados enviados com sucesso ao Google Sheets.");
		} catch (Exception e) {
			logger.log(Level.SEVERE, "Erro ao enviar dados ao Google Sheets", e);
			throw e;
		}
	}



	public static Sheets getSheetsService() throws Exception {
		String credentialsPath = dotenv.get("GOOGLE_CREDENTIALS_PATH");

		GoogleCredential credential = GoogleCredential.fromStream(new FileInputStream(credentialsPath))
				.createScoped(Collections.singleton(SheetsScopes.SPREADSHEETS));

		return new Sheets.Builder(credential.getTransport(), credential.getJsonFactory(), credential)
				.setApplicationName("Google Sheets API Java")
				.build();
	}

	private static void configurarLogger() {
		try {
			Logger rootLogger = Logger.getLogger("");
			Handler fileHandler = new FileHandler("smartbank.log", true);
			fileHandler.setFormatter(new SimpleFormatter());
			rootLogger.addHandler(fileHandler);
			rootLogger.setLevel(Level.INFO);
		} catch (Exception e) {
			System.err.println("Erro ao configurar o logger: " + e.getMessage());
		}
	}
}
