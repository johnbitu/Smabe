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
				readPlanilha(filePath, spreadsheetId);
				logger.info("Processamento concluído com sucesso!");
			} catch (Exception e) {
				logger.log(Level.SEVERE, "Erro durante o processamento", e);
			}
		} else {
			ExtractbankGUI.startGUI();
		}
	}

	public static void readPlanilha(String filePath, String spreadsheetId) throws Exception {
		logger.info("Lendo planilha: " + filePath);
		String passwordPj = dotenv.get("PASSWORD_PJ");
		BigDecimal entradaResult = BigDecimal.ZERO;
		BigDecimal saidaResult = BigDecimal.ZERO;
		BigDecimal sobraResult = BigDecimal.ZERO;

		try {
			FileInputStream fileIs = new FileInputStream(new File(filePath));
			POIFSFileSystem fileSystem = new POIFSFileSystem(fileIs);
			EncryptionInfo encryptionInfo = new EncryptionInfo(fileSystem);
			Decryptor decryptor = Decryptor.getInstance(encryptionInfo);

			if (!decryptor.verifyPassword(passwordPj)) {
				logger.severe("Senha incorreta para descriptografar a planilha.");
				throw new RuntimeException("Senha incorreta!");
			}

			try (Workbook workbook = new XSSFWorkbook(decryptor.getDataStream(fileSystem))) {
				Sheet sheet = workbook.getSheetAt(0);

				for (Row row : sheet) {
					if (row.getRowNum() == 0) continue;

					String titulo = (row.getCell(1) != null) ? row.getCell(1).toString() : "Sem título";
					String entrada = (row.getCell(3) != null) ? row.getCell(3).toString() : "0";
					String saida = (row.getCell(4) != null) ? row.getCell(4).toString() : "0";

					try {
						if (!entrada.trim().isEmpty()) {
							BigDecimal entradaNumber = new BigDecimal(entrada.replace(",", ".").trim());
							entradaResult = entradaResult.add(entradaNumber);
						}
						if (!saida.trim().isEmpty()) {
							BigDecimal saidaNumber = new BigDecimal(saida.replace(",", ".").trim());
							saidaResult = saidaResult.add(saidaNumber);
						}
					} catch (NumberFormatException e) {
						logger.warning("Erro ao processar a linha " + row.getRowNum() + ": Entrada = " + entrada + ", Saída = " + saida);
					}
				}

				saidaResult = saidaResult.multiply(BigDecimal.valueOf(-1));
				sobraResult = entradaResult.add(saidaResult);

				logger.info("Total Entrada: " + entradaResult + ", Total Saída: " + saidaResult + ", Saldo: " + sobraResult);
			}

			insertInGoogle(spreadsheetId, entradaResult, saidaResult, sobraResult);
		} catch (Exception e) {
			logger.log(Level.SEVERE, "Erro ao processar a planilha", e);
			throw e;
		}
	}

	public static void insertInGoogle(String spreadsheetId, BigDecimal entrada, BigDecimal saida, BigDecimal sobra) throws Exception {
		logger.info("Iniciando envio ao Google Sheets para ID: " + spreadsheetId);
		String range = "Dados1!A1:C2";
		Sheets sheetsService = getSheetsService();

		try {
			ValueRange body = new ValueRange().setValues(Arrays.asList(
					Arrays.asList("Entrada", "Saída", "Saldo"),
					Arrays.asList(entrada.toString(), saida.toString(), sobra.toString())
			));

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

