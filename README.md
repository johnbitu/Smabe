# Smartbank Extract

**Smartbank Extract** é uma aplicação Java desenvolvida para processar arquivos de dados bancários (como arquivos CSV e Excel) e enviar as informações organizadas para uma planilha do Google Sheets. O sistema suporta a leitura de arquivos criptografados e inclui uma interface gráfica (GUI) para facilitar o uso em ambientes não headless.

---

## **Funcionalidades**
- Processamento de arquivos CSV.
- Processamento de arquivos Excel (.xls e .xlsx), incluindo suporte a arquivos criptografados.
- Integração com Google Sheets para envio automatizado de dados.
- Execução em modos:
  - **Headless**: Execução via linha de comando.
  - **GUI**: Interface gráfica para seleção de arquivos e envio.

---

## **Tecnologias Utilizadas**
- **Linguagem:** Java 11+
- **Frameworks e Bibliotecas:**
  - Spring Boot
  - Apache POI (para manipulação de arquivos Excel)
  - Google Sheets API (para integração com o Google Sheets)
  - Dotenv (para gerenciamento de variáveis de ambiente)
- **APIs Externas:**
  - Google Sheets API

---

## **Instalação e Configuração**

### **1. Pré-requisitos**
- Java 11 ou superior instalado.
- Conta no Google Cloud com credenciais configuradas para acesso à API do Google Sheets.

### **2. Configuração do ambiente**
1. Crie um arquivo `.env` no diretório raiz do projeto com as seguintes variáveis:
    ```
    GOOGLE_CREDENTIALS_PATH=<caminho_para_credenciais_do_google>
    PASSWORD_PJ=<senha_para_planilhas_criptografadas>
    ```
2. Configure as credenciais do Google Sheets:
    - Habilite a API do Google Sheets em sua conta do Google Cloud.
    - Baixe o arquivo JSON das credenciais da API e salve no caminho definido em `GOOGLE_CREDENTIALS_PATH`.

### **3. Execução do projeto**
#### **Modo Headless**
Execute o projeto via terminal com os seguintes parâmetros:
```sh
java -jar smartbank-extract.jar <caminho-da-planilha> <id-da-planilha-google>
```
- **Exemplo**:
    ```sh
    java -jar smartbank-extract.jar "dados.xlsx" "1A2B3C4D5E6F"
    ```

#### **Modo GUI**
Basta executar o jar sem parâmetros:
```sh
java -jar smartbank-extract.jar
```
Uma interface gráfica será aberta, permitindo que você selecione os arquivos e o ID da planilha.

---

## **Estrutura do Projeto**

### **Pacotes principais**
- `com.dev.smartbankextract`:
  - **SmartbankextractApplication:** Classe principal que gerencia o fluxo do sistema.
  - **ExtractbankGUI:** Classe responsável por iniciar a interface gráfica.

### **Métodos principais**
- `readCsv`: Processa arquivos CSV e os prepara para envio ao Google Sheets.
- `readPlanilha`: Processa arquivos Excel (.xls ou .xlsx) e os prepara para envio ao Google Sheets.
- `insertInGoogle`: Envia os dados processados para o Google Sheets.
- `getSheetsService`: Configura e autentica o acesso à API do Google Sheets.
- `configurarLogger`: Configura o sistema de logs para registro de eventos.

---

## **Layouts de Planilha**
O sistema permite a customização do layout antes do envio dos dados. O layout padrão inclui:

| **Data**   | **Descrição** | **Categoria** | **Valor** | **Observações** |
|------------|---------------|---------------|-----------|-----------------|
| 01/01/2025 | Compra A      | Alimentação   | 100.00    |                 |
| 02/01/2025 | Compra B      | Transporte    | 50.00     |                 |

Você pode ajustar o layout no método `insertInGoogle` conforme suas necessidades.

---

## **Tratamento de Erros**
- **Erros na leitura de arquivos:** Logs são gerados caso o arquivo seja inválido ou a senha para descriptografia esteja incorreta.
- **Erros na API do Google Sheets:** O sistema tenta capturar e logar falhas na conexão ou envio de dados.

---

## **Logs**
Os logs são registrados em um arquivo `smartbank.log` no diretório atual e incluem informações sobre:
- Eventos bem-sucedidos, como leitura de arquivos e envio de dados.
- Erros ocorridos durante o processamento.

---

## **Possíveis Melhorias Futuras**
- Suporte a outros formatos de arquivo (como JSON ou XML).
- Funcionalidades de formatação avançada para o Google Sheets (cores, fórmulas automáticas, etc.).
- Implementação de um sistema de fila para envios assíncronos ao Google Sheets.

---

## **Contato**
Caso tenha dúvidas ou precise de suporte, entre em contato pelo e-mail: [bitujoaovictor@gmail.com](mailto:bitujoaovictor@gmail.com).
