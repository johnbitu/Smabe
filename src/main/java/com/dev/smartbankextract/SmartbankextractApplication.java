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

@SpringBootApplication
public class SmartbankextractApplication {

	private static final Dotenv dotenv = Dotenv.configure().load();

	public static void main(String[] args) {
		boolean isHeadless = GraphicsEnvironment.isHeadless();

		if (isHeadless) {
			System.out.println("Executando em ambiente headless...");
			if (args.length < 2) {
				System.out.println("Uso: java -jar sua-aplicacao.jar <caminho-da-planilha> <id-da-planilha-google>");
				return;
			}

			String filePath = args[0];
			String spreadsheetId = args[1];

			try {
				readPlanilha(filePath, spreadsheetId);
				System.out.println("Processamento concluído com sucesso!");
			} catch (Exception e) {
				System.err.println("Erro durante o processamento: " + e.getMessage());
				e.printStackTrace();
			}
		} else {
			ExtractbankGUI.startGUI();
		}
	}


	public static void readPlanilha(String filePath, String spreadsheetId) throws Exception {
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
						System.err.println("Erro ao processar a entrada na linha " + row.getRowNum() + ": " + entrada);
					}
				}

				saidaResult = saidaResult.multiply(BigDecimal.valueOf(-1));
				sobraResult = entradaResult.add(saidaResult);

				System.out.println("Total da Entrada: " + entradaResult);
				System.out.println("Total da Saída: " + saidaResult);
				System.out.println("Sobra: " + sobraResult);
			}

			insertInGoogle(spreadsheetId, entradaResult, saidaResult, sobraResult);
		} catch (Exception e) {
			throw new RuntimeException("Erro ao processar a planilha: " + e.getMessage(), e);
		}
	}

	public static void insertInGoogle(String spreadsheetId, BigDecimal entrada, BigDecimal saida, BigDecimal sobra) throws Exception {
		String range = "Dados1!A1:C2";
		Sheets sheetsService = getSheetsService();

		ValueRange body = new ValueRange().setValues(Arrays.asList(
				Arrays.asList("Entrada", "Saída", "Sobra"),
				Arrays.asList(entrada.toString(), saida.toString(), sobra.toString())
		));

		sheetsService.spreadsheets().values()
				.update(spreadsheetId, range, body)
				.setValueInputOption("RAW")
				.execute();

		System.out.println("Dados enviados com sucesso ao Google Sheets.");
	}

	public static Sheets getSheetsService() throws Exception {
		String credentialsPath = dotenv.get("GOOGLE_CREDENTIALS_PATH");

		GoogleCredential credential = GoogleCredential.fromStream(new FileInputStream(credentialsPath))
				.createScoped(Collections.singleton(SheetsScopes.SPREADSHEETS));

		return new Sheets.Builder(credential.getTransport(), credential.getJsonFactory(), credential)
				.setApplicationName("Google Sheets API Java")
				.build();
	}
}
