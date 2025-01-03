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

import java.io.File;
import java.io.FileInputStream;
import java.math.BigDecimal;
import java.util.Arrays;
import java.util.Collections;

@SpringBootApplication
public class SmartbankextractApplication {

	private static final Dotenv dotenv = Dotenv.configure().load();

	public static void main(String[] args) throws Exception {
		SpringApplication.run(SmartbankextractApplication.class, args);
		try {
			readPlanilha();
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	public static void readPlanilha() throws Exception {
		String password = dotenv.get("PASSWORD");
		String passwordPj = dotenv.get("PASSWORD_PJ");
		BigDecimal entradaResult = BigDecimal.ZERO;
		BigDecimal saidaResult = BigDecimal.ZERO;
		BigDecimal sobraResult = BigDecimal.ZERO;

		try {
			FileInputStream fileIs = new FileInputStream(new File("Extratos/extrato-dezembro.xlsx"));
			POIFSFileSystem fileSystem = new POIFSFileSystem(fileIs);
			EncryptionInfo encryptionInfo = new EncryptionInfo(fileSystem);
			Decryptor decryptor = Decryptor.getInstance(encryptionInfo);

			// Tentar desbloquear o arquivo com a senha
			if (!decryptor.verifyPassword(passwordPj)) {
				throw new RuntimeException("Senha incorreta!");
			}

			// Abrir o conteúdo descriptografado
			try (Workbook workbook = new XSSFWorkbook(decryptor.getDataStream(fileSystem))) {
				Sheet sheet = workbook.getSheetAt(0);

				for (Row row : sheet) {
					if (row.getRowNum() == 0) continue; // Ignora o cabeçalho

					// Verifica e extrai os valores das células
					String titulo = (row.getCell(1) != null) ? row.getCell(1).toString() : "Sem título";
					String entrada = (row.getCell(3) != null) ? row.getCell(3).toString() : "0";
					String saida = (row.getCell(4) != null) ? row.getCell(4).toString() : "0";

					try {
						if (!entrada.trim().isEmpty()) {
							String entradaLimpa = entrada.replace(",", ".").trim();
							BigDecimal entradaNumber = new BigDecimal(entradaLimpa);
							entradaResult = entradaResult.add(entradaNumber);
						}
						if (!saida.trim().isEmpty()) {
							String saidaLimpa = saida.replace(",", ".").trim();
							BigDecimal saidaNumber = new BigDecimal(saidaLimpa);
							saidaResult = saidaResult.add(saidaNumber);

						}
					} catch (NumberFormatException e) {
						System.err.println("Erro ao processar a entrada na linha " + row.getRowNum() + ": " + entrada);
					}

					System.out.println("Titulo/Descrição: " + titulo);
					System.out.println("Entrada: " + entrada);
					System.out.println("Saída: " + saida);
					System.out.println("--------------------------------------------------");
				}
				saidaResult = saidaResult.multiply(BigDecimal.valueOf(-1));
				sobraResult = entradaResult.add(saidaResult);

				System.out.println("Total da Entrada: " + entradaResult);
				System.out.println("Total da Saída: " + saidaResult);
				System.out.println("Sobra: " + sobraResult);
			}

			insertInGoogle(entradaResult, saidaResult, sobraResult);
		} catch (Exception e) {
			e.printStackTrace();
		}

	}

	public static void insertInGoogle(BigDecimal entrada, BigDecimal saida, BigDecimal sobra) throws Exception {
		String spreadsheetId = dotenv.get("SPREADSHEET_ID"); // Substitua pelo ID da sua planilha
		String range = "Dados1!A1:C2";
		Sheets sheetsService = getSheetsService();

		try {
			ValueRange body = new ValueRange().setValues(Arrays.asList(
					Arrays.asList("Entrada", "Saída", "Sobra"),
					Arrays.asList(entrada.toString(), saida.toString(), sobra.toString())
			));

			sheetsService.spreadsheets().values()
					.update(spreadsheetId, range, body)
					.setValueInputOption("RAW")
					.execute();

			System.out.println("Dados enviados com sucesso ao Google Sheets.");
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	public static Sheets getSheetsService() throws Exception {
		String credentialsPath = dotenv.get("GOOGLE_CREDENTIALS_PATH"); // Substitua pelo caminho real do arquivo de credenciais

		GoogleCredential credential = GoogleCredential.fromStream(new FileInputStream(credentialsPath))
				.createScoped(Collections.singleton(SheetsScopes.SPREADSHEETS));

		return new Sheets.Builder(credential.getTransport(), credential.getJsonFactory(), credential)
				.setApplicationName("Google Sheets API Java")
				.build();
	}
}
