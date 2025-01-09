package com.dev.smartbankextract;

import com.google.api.client.googleapis.auth.oauth2.GoogleCredential;
import com.google.api.services.sheets.v4.Sheets;
import com.google.api.services.sheets.v4.SheetsScopes;
import com.google.api.services.sheets.v4.model.*;
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
import java.text.DecimalFormat;
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
	private static final DecimalFormat decimalFormat = new DecimalFormat("0.00");



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

		boolean isFirstLine = true;

		try (BufferedReader br = new BufferedReader(new FileReader(filePath))) {
			String line;
			while ((line = br.readLine()) != null) {

				if (isFirstLine) {
					isFirstLine = false; // Ignora a primeira linha e desativa a flag
					continue;
				}

				String[] values = line.split(","); // Divide a linha usando vírgulas como delimitador

				// Processa a descrição usando regex com "-" como delimitador
				String descricaoCompleta = values.length > 3 ? values[3] : "";
				String[] descricaoPartes = descricaoCompleta.split(" - ", 3); // Divide em até 3 partes
				String titulo = descricaoPartes.length > 1 ? descricaoPartes[0] + " " + descricaoPartes[1] : descricaoCompleta;
				String descricao = descricaoPartes.length > 0 ? descricaoPartes[0] : "";

				// Adiciona os valores no formato do layout desejado
				List<Object> linha = new ArrayList<>();
				linha.add(values.length > 0 ? values[0] : ""); // Data
				linha.add(titulo); // Título gerado
				linha.add(descricao); // Descrição gerada

				// Diferencia entradas e saídas com base no valor
				BigDecimal valor = values.length > 1 && !values[1].isEmpty() ? new BigDecimal(values[1]) : BigDecimal.ZERO;
//				String valorFormatado = decimalFormat.format(valor);
				linha.add(valor.compareTo(BigDecimal.ZERO) > 0 ? valor : decimalFormat.format(BigDecimal.ZERO)); // Entrada
				linha.add(valor.compareTo(BigDecimal.ZERO) < 0 ? valor : decimalFormat.format(BigDecimal.ZERO)); // Saída

				linha.add(""); // Categoria (vazio conforme especificado)
				linha.add(""); // Observações (vazio conforme especificado)

				linha.add("NuBank"); // Banco

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

							linha.add(""); // Categoria
							linha.add(""); // Observações

							linha.add("C6 Bank");

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
		Sheets sheetsService = getSheetsService();

		String sheetName = "SmabeDados";

		// Verifica se a aba 'SmabeDados' existe
		Integer sheetId = null;
		try {
			sheetId = getSheetIdByName(sheetsService, spreadsheetId, sheetName);
			logger.info("Aba '" + sheetName + "' encontrada.");
		} catch (RuntimeException e) {
			logger.info("Aba '" + sheetName + "' não encontrada. Criando uma nova.");
			createSheet(sheetsService, spreadsheetId, sheetName);
			sheetId = getSheetIdByName(sheetsService, spreadsheetId, sheetName);
		}

		// Determina a última linha preenchida
		String rangeToGetLastRow = sheetName + "!A:A"; // Supondo que os dados estão na aba "Dados1" na coluna A
		ValueRange response = sheetsService.spreadsheets().values()
				.get(spreadsheetId, rangeToGetLastRow)
				.execute();

		List<List<Object>> values = response.getValues();
		int lastRow = values != null ? values.size() : 0; // Conta o número de linhas preenchidas

		logger.info("Última linha preenchida: " + lastRow);

		// Define o intervalo inicial para adicionar os novos dados
		String range = sheetName + "!A" + (lastRow + 1); // Começa a partir da próxima linha disponível

		// Monta o layout da planilha
		List<List<Object>> planilhaComLayout = new ArrayList<>();

		// Adiciona o cabeçalho apenas se a planilha estiver vazia
		if (lastRow == 0) {
			List<Object> cabecalho = Arrays.asList("Data", "Titulo", "Descrição", "Entrada", "Saída", "Categoria", "Observações", "Banco");
			planilhaComLayout.add(cabecalho);
		}

		// Adiciona os dados lidos no layout
		for (List<Object> registro : registros) {
			List<Object> linhaFormatada = new ArrayList<>();

			// Organiza os valores de Entrada e Saída
			String entrada = registro.size() > 3 ? registro.get(3).toString() : "";
			String saida = registro.size() > 4 ? registro.get(4).toString() : "";
			linhaFormatada.add(registro.size() > 0 ? registro.get(0) : ""); // Data
			linhaFormatada.add(registro.size() > 1 ? registro.get(1) : ""); // Titulo
			linhaFormatada.add(registro.size() > 2 ? registro.get(2) : ""); // Descrição
			linhaFormatada.add(!entrada.isEmpty() ? entrada : ""); // Entrada
			linhaFormatada.add(!saida.isEmpty() ? saida : ""); // Saída
			linhaFormatada.add(""); // Categoria (vazio)
			linhaFormatada.add(""); // Observações (vazio)
			linhaFormatada.add(registro.size() > 7 ? registro.get(7) : ""); // Banco

			planilhaComLayout.add(linhaFormatada);
		}

		// Envia os dados ao Google Sheets
		try {
			ValueRange body = new ValueRange().setValues(planilhaComLayout);

			sheetsService.spreadsheets().values()
					.append(spreadsheetId, range, body) // Usando 'append' para adicionar linhas
					.setValueInputOption("RAW")
					.execute();

			logger.info("Dados enviados com sucesso ao Google Sheets.");

			ordenarDadosPorColuna(sheetsService, spreadsheetId, "Dados1", 0, "DESCENDING");
		} catch (Exception e) {
			logger.log(Level.SEVERE, "Erro ao enviar dados ao Google Sheets", e);
			throw e;
		}

	}

	public static void createSheet(Sheets sheetsService, String spreadsheetId, String sheetName) throws Exception {
		try {
			BatchUpdateSpreadsheetRequest requestBody = new BatchUpdateSpreadsheetRequest()
					.setRequests(Collections.singletonList(
							new Request().setAddSheet(new AddSheetRequest()
									.setProperties(new SheetProperties().setTitle(sheetName))
							)
					));

			sheetsService.spreadsheets().batchUpdate(spreadsheetId, requestBody).execute();
			logger.info("Aba '" + sheetName + "' criada com sucesso.");
		} catch (Exception e) {
			logger.log(Level.SEVERE, "Erro ao criar aba no Google Sheets", e);
			throw e;
		}
	}

	public static void ordenarDadosPorColuna(Sheets sheetsService, String spreadsheetId, String sheetName, int columnIndex, String sortOrder) throws Exception {
		try {
			Integer sheetId = getSheetIdByName(sheetsService, spreadsheetId, sheetName);

			GridRange sortRange = new GridRange()
					.setSheetId(sheetId)
					.setStartRowIndex(1)
					.setStartColumnIndex(0)
					.setEndColumnIndex(7);

			SortRangeRequest sortRequest = new SortRangeRequest()
					.setRange(sortRange)
					.setSortSpecs(Collections.singletonList(
							new SortSpec()
									.setDimensionIndex(columnIndex)
									.setSortOrder(sortOrder)
					));

			Request request = new Request().setSortRange(sortRequest);
			BatchUpdateSpreadsheetRequest batchRequest = new BatchUpdateSpreadsheetRequest()
					.setRequests(Collections.singletonList(request));

			sheetsService.spreadsheets().batchUpdate(spreadsheetId, batchRequest).execute();
			logger.info("Coluna ordenada com sucesso em ordem " + sortOrder + ".");
		} catch (Exception e) {
			logger.log(Level.SEVERE, "Erro ao ordenar dados na planilha", e);
			throw e;
		}
	}

	public static Integer getSheetIdByName(Sheets sheetsService, String spreadsheetId, String sheetName) throws Exception {
		try {
			Spreadsheet spreadsheet = sheetsService.spreadsheets().get(spreadsheetId).execute();
			return spreadsheet.getSheets().stream()
					.filter(sheet -> sheet.getProperties().getTitle().equals(sheetName))
					.map(sheet -> sheet.getProperties().getSheetId())
					.findFirst()
					.orElseThrow(() -> new RuntimeException("Aba com o nome '" + sheetName + "' não encontrada."));
		} catch (Exception e) {
			logger.log(Level.SEVERE, "Erro ao obter o ID da aba", e);
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
