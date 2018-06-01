package br.com.pangare.gerador;
import java.io.File;
import java.io.FileOutputStream;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.ss.util.WorkbookUtil;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class GravaXLS {

	private final static String VERSUS = "x";
	private final static char AD = '"'; // Aspas duplas
	private final static int COLUNA_ZERO = 0;

	// Sheet Jogos
	private static final int JOGO_COL_RODADA = 0;
	private static final int JOGO_COL_JOGO = 1;
	private static final int JOGO_COL_TIME1 = 2;
	private static final int JOGO_COL_GOLS_TIME1 = 3;
	private static final int JOGO_COL_VERSUS = 4;
	private static final int JOGO_COL_GOLS_TIME2 = 5;
	private static final int JOGO_COL_TIME2 = 6;
	private static final int JOGO_COL_ARBITRAGEM = 7;
	private static final int JOGO_COL_AUX1 = 8;
	private static final int JOGO_COL_AUX2 = 9;

	// Sheet classificacao
	private static final int CLASS_COL_CLASS = 0;
	private static final int CLASS_COL_TIME = 1;
	private static final int CLASS_COL_PONTO_GANHO = 2;
	private static final int CLASS_COL_PONTO_PERDIDO = 3;
	private static final int CLASS_COL_JOGO = 4;
	private static final int CLASS_COL_VITORIA = 5;
	private static final int CLASS_COL_EMPATE = 6;
	private static final int CLASS_COL_DERROTA = 7;
	private static final int CLASS_COL_GOL_PRO = 8;
	private static final int CLASS_COL_GOL_CONTRA = 9;
	private static final int CLASS_COL_SALDO = 10;
	private static final int CLASS_COL_PORCENTAGEM = 11;
	private static final int CLASS_COL_ARBITRAGEM = 12;
	
	private static XSSFCellStyle styleCabecalho;
	private static XSSFCellStyle styleCabecalhoCenter;
	private static XSSFCellStyle styleCabecalhoLeft;
	private static XSSFCellStyle styleCabecalhoRight;
	private static XSSFCellStyle styleTabela;
	private static XSSFCellStyle styleTabelaCenter;
	private static XSSFCellStyle styleTabelaLeft;
	private static XSSFCellStyle styleTabelaRight;
	private static XSSFCellStyle styleTabelaPorcentagem;
	private static XSSFCellStyle styleCorpo;
	private static XSSFCellStyle styleCorpoRight;
	private static XSSFCellStyle styleCorpoLeft;
	private static XSSFCellStyle styleCorpoCenter;
	private static XSSFCellStyle styleCorpoPorcentagem;
	private static XSSFCellStyle styleCorpoMedia;
	private static XSSFCellStyle styleTitulo;
	private static XSSFCellStyle styleFase;
	private static XSSFCellStyle styleTabelaPontoPerdido;
	private static XSSFCellStyle styleDestaqueCenter;
	
	private static int linhaInicialJogos = 0;
	private static int linhaFinalJogos = 0;

	CreationHelper createHelper;

	
	public static void criaSheetJogos(XSSFWorkbook wb, Campeonato campeonato) {
		
	    // cria sheet para jogos
	    String safeName = WorkbookUtil.createSafeSheetName("Jogos"); 
	    XSSFSheet sheet = (XSSFSheet) wb.createSheet(safeName);
	    CreationHelper createHelper = sheet.getWorkbook().getCreationHelper();
	    
	    configuraStyles(wb, campeonato);
	    
	    // Largura das colunas
	    sheet.setColumnWidth(JOGO_COL_RODADA, 256 * 8);
	    sheet.setColumnWidth(JOGO_COL_JOGO, 256 * 8);
	    sheet.setColumnWidth(JOGO_COL_TIME1, 256 * 22);
	    sheet.setColumnWidth(JOGO_COL_GOLS_TIME1,256 * 4);
	    sheet.setColumnWidth(JOGO_COL_VERSUS,256 * 3);
	    sheet.setColumnWidth(JOGO_COL_GOLS_TIME2, 256 * 4);
	    sheet.setColumnWidth(JOGO_COL_TIME2,256 * 22);
	    if (campeonato.getIsArbitragem())
	    	sheet.setColumnWidth(JOGO_COL_ARBITRAGEM, 256 * 14);
	    
	    // hide colunas auxiliares
	    sheet.setColumnHidden(JOGO_COL_AUX1, true);
	    sheet.setColumnHidden(JOGO_COL_AUX2, true);

	    int linhaAtual = 0;
	    
		// Escreve Titulo
	    Row row = sheet.createRow(linhaAtual);
	    escreveCell(row, COLUNA_ZERO, styleTitulo, campeonato.getNomeCampeonato());
	    
		// Escreve fase
	    linhaAtual = linhaAtual + 2;
	    row = sheet.createRow(linhaAtual);
	    escreveCell(row, COLUNA_ZERO , styleFase, "Fase classificatória");

	    linhaAtual++;
    	linhaInicialJogos = linhaAtual;;

	    
	    for (Grupo grupo : campeonato.getGrupos()) {

			// Escreve nome do Grupo (se houver mais de um grupo)
	    	if (campeonato.getNumeroGrupos() > 1) {
	    		linhaAtual++;
	    		row = sheet.createRow(linhaAtual);
	    		escreveCell(row, COLUNA_ZERO , styleFase, grupo.getNomeGrupo());
	    	}
	    	
	    	// Escreve Cabecalho

		    linhaAtual++;
		    row = sheet.createRow(linhaAtual);
	    	escreveCell(row, JOGO_COL_RODADA, styleCabecalhoCenter, "Rodada");
	    	escreveCell(row, JOGO_COL_JOGO, styleCabecalhoCenter, "Jogo");
	    	escreveCell(row, JOGO_COL_TIME1, styleCabecalhoLeft, "Time mandante");
	    	escreveCell(row, JOGO_COL_GOLS_TIME1, styleCabecalhoCenter, "");
	    	escreveCell(row, JOGO_COL_VERSUS, styleCabecalhoCenter, "x");
	    	escreveCell(row, JOGO_COL_GOLS_TIME2, styleCabecalhoCenter, "");
	    	escreveCell(row, JOGO_COL_TIME2, styleCabecalhoRight, "Time visitante");
	    	if (campeonato.getIsArbitragem())
	    	escreveCell(row, JOGO_COL_ARBITRAGEM, styleCabecalhoLeft, "Arbitragem");

	    	// Escreve jogos
	    	linhaAtual++;

	    	String formula;
	    	CellReference crGols1, crGols2;
    		
	    	for (Jogo jogo : grupo.getTabelaJogos()) {

	    		row = sheet.createRow(linhaAtual++);
	    	
	    		escreveCell(row, JOGO_COL_RODADA, styleTabelaCenter, (double) jogo.getRodada());

	    		escreveCell(row, JOGO_COL_JOGO, styleTabelaCenter, (double) jogo.getNumJogo());
	    	
	    		escreveCell(row, JOGO_COL_TIME1, styleTabelaLeft, String.valueOf(jogo.getTimeMandante()));
	    	
	    		if (campeonato.getIsValoresAleatorios())
	    			escreveCell(row, JOGO_COL_GOLS_TIME1, styleTabelaCenter, (double) (int) (Math.random() * 5));
	    		else
	    			escreveCell(row, JOGO_COL_GOLS_TIME1, styleTabelaCenter, "");
	    		crGols1 = new CellReference(row.getRowNum(), JOGO_COL_GOLS_TIME1);

	    		escreveCell(row, JOGO_COL_VERSUS, styleTabelaCenter,VERSUS);
	    	
	    		if (campeonato.getIsValoresAleatorios())
	    			escreveCell(row, JOGO_COL_GOLS_TIME2, styleTabelaCenter, (double) (int) (Math.random() * 5));
	    		else
	    			escreveCell(row, JOGO_COL_GOLS_TIME2, styleTabelaCenter, "");
	    		crGols2 = new CellReference(row.getRowNum(), JOGO_COL_GOLS_TIME2);

	    	
	    		escreveCell(row, JOGO_COL_TIME2, styleTabelaRight, String.valueOf(jogo.getTimeVisitante()));

	    		if (campeonato.getIsArbitragem()) {
	    			if (campeonato.getIsValoresAleatorios())
	    				escreveCell(row, JOGO_COL_ARBITRAGEM, styleTabelaLeft, grupo.getTimes().get( (int) (Math.random() * grupo.getTimes().size() - 1)));
	    			else
	    				escreveCell(row, JOGO_COL_ARBITRAGEM, styleTabelaLeft, "");
	    		}	    		
	    		
	    	
	    		// Coluna auxiliar 1
	    		formula = "IF(OR(" + crGols1.formatAsString() + "=" + AD + AD + "," + crGols2.formatAsString() + "=" + AD + AD + ")," + AD + AD + ",IF(" + crGols1.formatAsString() + ">" + crGols2.formatAsString() +"," + AD + "V" + AD + ",IF(" + crGols1.formatAsString() + "="+ crGols2.formatAsString() + "," + AD + "E" + AD + ",IF(" + crGols1.formatAsString() + "<" + crGols2.formatAsString() +"," + AD + "D" + AD + "))))";
	    		escreveFormula(row, JOGO_COL_AUX1, styleTabelaCenter, formula);

	    		// Coluna auxiliar 2
	    		formula = "IF(OR(" + crGols2.formatAsString() + "=" + AD + AD + "," + crGols1.formatAsString() + "=" + AD + AD + ")," + AD + AD + ",IF(" + crGols2.formatAsString() + ">" + crGols1.formatAsString() +"," + AD + "V" + AD + ",IF(" + crGols2.formatAsString() + "="+ crGols1.formatAsString() + "," + AD + "E" + AD + ",IF(" + crGols2.formatAsString() + "<" + crGols1.formatAsString() +"," + AD + "D" + AD + "))))";
	    		escreveFormula(row, JOGO_COL_AUX2, styleTabelaCenter, formula);

	    	}
	    
	    }
	    
    	linhaFinalJogos = linhaAtual;

    	/*
    	
		// Escreve fase final
    	linhaAtual++;
	    row = sheet.createRow(linhaAtual);
    	escreveCell(row, COLUNA_ZERO, styleFase, "Finais");

    	// escreveCabecalho
    	linhaAtual++;
    	linhaAtual++;
	    row = sheet.createRow(linhaAtual);
    	sheet.copyRows(JOGO_LINHA_CABECALHO, JOGO_LINHA_CABECALHO, linhaAtual, new CellCopyPolicy());
    	
    	// partidas finais

    	int numJogo = campeonato.getTabelaJogos().size() + 1;
    	
    	linhaJogo++;
	    row = sheet.createRow(linhaJogo);
	    escreveCell(row, JOGO_COL_RODADA, styleTabelaCenter, "1");
	    escreveCell(row, JOGO_COL_JOGO, styleTabelaCenter, (double) numJogo++);
	    escreveFormula(row, JOGO_COL_TIME1, styleTabelaLeft, "Classifica��o!B7");
	    escreveCell(row, JOGO_COL_GOLS_TIME1, styleTabelaCenter, "");
	    escreveCell(row, JOGO_COL_VERSUS, styleTabelaCenter, VERSUS);
	    escreveCell(row, JOGO_COL_GOLS_TIME2, styleTabelaCenter, "");
	    escreveFormula(row, JOGO_COL_TIME2, styleTabelaRight, "Classifica��o!B6");
	    if (campeonato.getIsArbitragem())
	    	escreveCell(row, JOGO_COL_ARBITRAGEM, styleTabelaCenter, "");

	    row = sheet.createRow(++linhaJogo);
	    escreveCell(row, JOGO_COL_RODADA, styleTabelaCenter, "2");
	    escreveCell(row, JOGO_COL_JOGO, styleTabelaCenter, (double) numJogo++);
	    escreveFormula(row, JOGO_COL_TIME1, styleTabelaLeft, "Classifica��o!B6");
	    escreveCell(row, JOGO_COL_GOLS_TIME1, styleTabelaCenter, "");
	    escreveCell(row, JOGO_COL_VERSUS, styleTabelaCenter, VERSUS);
	    escreveCell(row, JOGO_COL_GOLS_TIME2, styleTabelaCenter, "");
	    escreveFormula(row, JOGO_COL_TIME2, styleTabelaRight, "Classifica��o!B7");
	    if (campeonato.getIsArbitragem())
	    	escreveCell(row, JOGO_COL_ARBITRAGEM, styleTabelaRight, "");

	*/

	}

	public static void criaSheetAuxiliar(XSSFWorkbook wb, Campeonato campeonato) {

		// cria sheet auxiliar
	    String safeName = WorkbookUtil.createSafeSheetName("Auxiliar"); 
	    XSSFSheet sheet = (XSSFSheet) wb.createSheet(safeName);
	    	
	    int linhaAtual = 0;
	    
	    // Escreve titulo
	    Row row = sheet.createRow(linhaAtual++);
	    escreveCell(row, COLUNA_ZERO, styleFase, "Auxiliar");

	    String intervaloTimes1 = retornaColuna(JOGO_COL_TIME1) + String.valueOf(linhaInicialJogos) + ":" + retornaColuna(JOGO_COL_TIME1) + String.valueOf(linhaFinalJogos);
	    String intervaloTimes2 = retornaColuna(JOGO_COL_TIME2) + String.valueOf(linhaInicialJogos) + ":" + retornaColuna(JOGO_COL_TIME2) + String.valueOf(linhaFinalJogos);
	    String intervaloPlacar1 = retornaColuna(JOGO_COL_GOLS_TIME1) + String.valueOf(linhaInicialJogos) + ":" + retornaColuna(JOGO_COL_GOLS_TIME1) + String.valueOf(linhaFinalJogos);
	    String intervaloPlacar2 = retornaColuna(JOGO_COL_GOLS_TIME2) + String.valueOf(linhaInicialJogos) + ":" + retornaColuna(JOGO_COL_GOLS_TIME2) + String.valueOf(linhaFinalJogos);
	    String intervaloResultado1 = retornaColuna(JOGO_COL_AUX1) + String.valueOf(linhaInicialJogos) + ":" + retornaColuna(JOGO_COL_AUX1) + String.valueOf(linhaFinalJogos);
	    String intervaloResultado2 = retornaColuna(JOGO_COL_AUX2) + String.valueOf(linhaInicialJogos) + ":" + retornaColuna(JOGO_COL_AUX2) + String.valueOf(linhaFinalJogos);
	    String intervaloArbitragem = retornaColuna(JOGO_COL_ARBITRAGEM) + String.valueOf(linhaInicialJogos) + ":" + retornaColuna(JOGO_COL_ARBITRAGEM) + String.valueOf(linhaFinalJogos);
	    
    	CellReference crTime, crVitoria, crEmpate, crDerrota, crSaldo, crJogo, crPontosGanhos, crGolsPro, crGolsContra;
    	String formula;

    	// Escreve Cabecalho
    	row = sheet.createRow(linhaAtual++);
    	escreveCell(row, CLASS_COL_CLASS, styleCabecalhoCenter, "Pos");
    	escreveCell(row, CLASS_COL_TIME, styleCabecalhoCenter, "Time");
    	escreveCell(row, CLASS_COL_PONTO_GANHO, styleCabecalhoCenter, "PG");
    	escreveCell(row, CLASS_COL_PONTO_PERDIDO, styleCabecalhoCenter, "PP");
    	escreveCell(row, CLASS_COL_JOGO, styleCabecalhoCenter, "J");
    	escreveCell(row, CLASS_COL_VITORIA, styleCabecalhoCenter, "V");
    	escreveCell(row, CLASS_COL_EMPATE, styleCabecalhoCenter, "E");
    	escreveCell(row, CLASS_COL_DERROTA, styleCabecalhoCenter, "D");
    	escreveCell(row, CLASS_COL_GOL_PRO, styleCabecalhoCenter, "GP");
    	escreveCell(row, CLASS_COL_GOL_CONTRA, styleCabecalhoCenter, "GC");
    	escreveCell(row, CLASS_COL_SALDO, styleCabecalhoCenter, "S");
    	escreveCell(row, CLASS_COL_PORCENTAGEM, styleCabecalhoRight, "Porc");
    	if (campeonato.getIsArbitragem())
    		escreveCell(row, CLASS_COL_ARBITRAGEM, styleCabecalhoCenter, "Arb");
    	
	    for (Grupo grupo : campeonato.getGrupos()) {

	    	// referencia de inicio das linhas de grupo
	    	grupo.setReferenciaInicialAux(new CellReference(row.getRowNum() + 1, CLASS_COL_CLASS));
	    	
	    	for (String time : grupo.getTimes()) {

	    		row = sheet.createRow(linhaAtual++);


	    		// Time
	    		escreveCell(row, CLASS_COL_TIME, styleTabelaLeft, time);
	    		crTime = new CellReference(row.getRowNum(), CLASS_COL_TIME);

	    		// Vitoria
	    		formula = "COUNTIFS(Jogos!" + intervaloTimes1 + "," + crTime.formatAsString() + ",Jogos!" + intervaloResultado1 + "," + AD + "V" + AD + ") + COUNTIFS(Jogos!" + intervaloTimes2 + "," + crTime.formatAsString() + ",Jogos!" + intervaloResultado2 + "," + AD + "V" + AD + ")";
	    		escreveFormula(row, CLASS_COL_VITORIA, styleTabelaCenter, formula);
	    		crVitoria = new CellReference(row.getRowNum(), CLASS_COL_VITORIA);

	    		// Empate
	    		formula = "COUNTIFS(Jogos!" + intervaloTimes1 + "," + crTime.formatAsString() + ",Jogos!" + intervaloResultado1 + "," + AD + "E" + AD + ") + COUNTIFS(Jogos!" + intervaloTimes2 + "," + crTime.formatAsString() + ",Jogos!" + intervaloResultado2 + "," + AD + "E" + AD + ")";
	    		escreveFormula(row, CLASS_COL_EMPATE, styleTabelaCenter, formula);
	    		crEmpate = new CellReference(row.getRowNum(), CLASS_COL_EMPATE);

	    		// Derrota
	    		formula = "COUNTIFS(Jogos!" + intervaloTimes1 + "," + crTime.formatAsString() + ",Jogos!" + intervaloResultado1 + "," + AD + "D" + AD + ") + COUNTIFS(Jogos!" + intervaloTimes2 + "," + crTime.formatAsString() + ",Jogos!" + intervaloResultado2 + "," + AD + "D" + AD + ")";
	    		escreveFormula(row, CLASS_COL_DERROTA, styleTabelaCenter, formula);
	    		crDerrota = new CellReference(row.getRowNum(), CLASS_COL_DERROTA);

	    		// Jogo
	    		formula = crVitoria.formatAsString() + "+" + crEmpate.formatAsString() + "+" + crDerrota.formatAsString(); 
	    		escreveFormula(row, CLASS_COL_JOGO, styleTabelaCenter, formula);
	    		crJogo = new CellReference(row.getRowNum(), CLASS_COL_JOGO);

	    		// PG
	    		formula = crVitoria.formatAsString() + "*3+" + crEmpate.formatAsString(); 
	    		escreveFormula(row, CLASS_COL_PONTO_GANHO, styleTabelaCenter,formula);
	    		crPontosGanhos = new CellReference(row.getRowNum(), CLASS_COL_PONTO_GANHO);

	    		// PP
	    		formula = crDerrota.formatAsString() + "*3+" + crEmpate.formatAsString() + "*2"; 
	    		escreveFormula(row, CLASS_COL_PONTO_PERDIDO, styleTabelaPontoPerdido, formula);
	    	
	    		// GP
	    		formula = "SUMIFS(Jogos!" + intervaloPlacar1 + ",Jogos!" + intervaloTimes1 + "," + crTime.formatAsString() + ") + SUMIFS(Jogos!" + intervaloPlacar2 + ",Jogos!" + intervaloTimes2 + ", " + crTime.formatAsString() + ")"; 
	    		escreveFormula(row, CLASS_COL_GOL_PRO, styleTabelaCenter, formula);
	    		crGolsPro = new CellReference(row.getRowNum(), CLASS_COL_GOL_PRO);

	    		// GC
	    		formula = "SUMIFS(Jogos!" + intervaloPlacar2 + ",Jogos!" + intervaloTimes1 + "," + crTime.formatAsString() + ") + SUMIFS(Jogos!" + intervaloPlacar1 + ",Jogos!" + intervaloTimes2 + ", " + crTime.formatAsString() + ")"; 
	    		escreveFormula(row, CLASS_COL_GOL_CONTRA, styleTabelaCenter, formula);
	    		crGolsContra = new CellReference(row.getRowNum(), CLASS_COL_GOL_CONTRA);

	    		// S
	    		formula = crGolsPro.formatAsString() + "-" + crGolsContra.formatAsString();
	    		escreveFormula(row, CLASS_COL_SALDO, styleTabelaCenter, formula);
	    		crSaldo = new CellReference(row.getRowNum(), CLASS_COL_SALDO);

	    		// %
	    		formula = "IF(" + crJogo.formatAsString() + ">0," + crPontosGanhos.formatAsString() + "/(" + crJogo.formatAsString() + "*3)," + AD + AD + ")";
	    		escreveFormula(row, CLASS_COL_PORCENTAGEM, styleTabelaPorcentagem, formula);

	    		// Arbitragens
	    		if (campeonato.getIsArbitragem()) {
	    			formula = "COUNTIF(Jogos!" + intervaloArbitragem + "," + crTime.formatAsString() + ")";
	    			escreveFormula(row, CLASS_COL_ARBITRAGEM, styleTabelaCenter, formula);
	    		}
	    
	    		// Pontuacao para classificacao automatica
	    		// =(D7*100)+(G7*50)+(L7*10)+J7+ROW()/100
	    		formula = crPontosGanhos.formatAsString() + "*100+" + crVitoria.formatAsString() + "*50+" + crSaldo.formatAsString() + "*10+" + crGolsPro.formatAsString() + "+ROW()/10";
	    		escreveFormula(row, CLASS_COL_CLASS, styleTabelaCenter, formula);

	    	}

	    	// referencia de inicio das linhas de grupo
    		if (campeonato.getIsArbitragem()) 
    			grupo.setReferenciaFinalAux(new CellReference(row.getRowNum(), CLASS_COL_ARBITRAGEM));
    		else
    			grupo.setReferenciaFinalAux(new CellReference(row.getRowNum(), CLASS_COL_PORCENTAGEM));
    	
    		grupo.setReferenciaUltimaLinhaAux(new CellReference(row.getRowNum(), CLASS_COL_CLASS));
	
	    }

	}
	
	public static void criaSheetClassificacao(XSSFWorkbook wb, Campeonato campeonato) {
	
		// cria sheet para classificação
	    String safeName = WorkbookUtil.createSafeSheetName("Classificação"); 
	    XSSFSheet sheet = (XSSFSheet) wb.createSheet(safeName);

	    // Largura das colunas
	    sheet.setColumnWidth(CLASS_COL_CLASS, 256 * 5);  
	    sheet.setColumnWidth(CLASS_COL_TIME, 256 * 20); 
	    sheet.setColumnWidth(CLASS_COL_PONTO_GANHO, 256 * 8);  
	    sheet.setColumnWidth(CLASS_COL_PONTO_PERDIDO, 256 * 8);  
	    sheet.setColumnWidth(CLASS_COL_JOGO, 256 * 5);  
	    sheet.setColumnWidth(CLASS_COL_VITORIA, 256 * 5);
	    sheet.setColumnWidth(CLASS_COL_EMPATE, 256 * 5);
 	    sheet.setColumnWidth(CLASS_COL_DERROTA, 256 * 5);
	    sheet.setColumnWidth(CLASS_COL_GOL_PRO, 256 * 5);
	    sheet.setColumnWidth(CLASS_COL_GOL_CONTRA, 256 * 5);
	    sheet.setColumnWidth(CLASS_COL_SALDO, 256 * 5);
	    sheet.setColumnWidth(CLASS_COL_PORCENTAGEM, 256 * 8);
	    if (campeonato.getIsArbitragem())
	    	sheet.setColumnWidth(CLASS_COL_ARBITRAGEM, 256 * 5);

	    int linhaAtual = 0;
	    
	    // Escreve titulo
	    Row row = sheet.createRow(linhaAtual);
	    escreveCell(row, COLUNA_ZERO, styleTitulo, campeonato.getNomeCampeonato());

		// Escreve fase
	    linhaAtual = linhaAtual + 2;
	    row = sheet.createRow(linhaAtual);
	    escreveCell(row, COLUNA_ZERO, styleFase, "Classificação 1a. fase");

    	String formula;

	    linhaAtual++;
    	int linhaInicial = linhaAtual;
    	
	    for (Grupo grupo : campeonato.getGrupos()) {

		    linhaAtual++;

			// Escreve nome do Grupo (se houver mais de um grupo)
	    	if (campeonato.getNumeroGrupos() > 1) {
	    		row = sheet.createRow(linhaAtual);
	    		escreveCell(row, COLUNA_ZERO, styleFase, grupo.getNomeGrupo());
	    	}
	    	
	    	// Escreve Cabecalho
	    	row = sheet.createRow(linhaAtual++);
	    	escreveCell(row, CLASS_COL_CLASS, styleCabecalhoCenter, "Pos");
	    	escreveCell(row, CLASS_COL_TIME, styleCabecalhoLeft, "Time");
	    	escreveCell(row, CLASS_COL_PONTO_GANHO, styleCabecalhoCenter, "PG");
	    	escreveCell(row, CLASS_COL_PONTO_PERDIDO, styleCabecalhoCenter, "PP");
	    	escreveCell(row, CLASS_COL_JOGO, styleCabecalhoCenter, "J");
	    	escreveCell(row, CLASS_COL_VITORIA, styleCabecalhoCenter, "V");
	    	escreveCell(row, CLASS_COL_EMPATE, styleCabecalhoCenter, "E");
	    	escreveCell(row, CLASS_COL_DERROTA, styleCabecalhoCenter, "D");
	    	escreveCell(row, CLASS_COL_GOL_PRO, styleCabecalhoCenter, "GP");
	    	escreveCell(row, CLASS_COL_GOL_CONTRA, styleCabecalhoCenter, "GC");
	    	escreveCell(row, CLASS_COL_SALDO, styleCabecalhoCenter, "S");
	    	escreveCell(row, CLASS_COL_PORCENTAGEM, styleCabecalhoRight, "Porc");
	    	if (campeonato.getIsArbitragem())
	    		escreveCell(row, CLASS_COL_ARBITRAGEM, styleCabecalhoCenter, "Arb");
	    
	    	String intervaloMaior = "Auxiliar!" + grupo.getReferenciaInicialAux().formatAsString() + ":" + grupo.getReferenciaUltimaLinhaAux().formatAsString();
	    	String intervaloGrupo = "Auxiliar!" +  grupo.getReferenciaInicialAux().formatAsString()+ ":" + grupo.getReferenciaFinalAux().formatAsString();
	    	
	    	CellReference crJogo;
	    	
	    	for (int i = 1; i < grupo.getTimes().size() + 1; i++) {

	    		row = sheet.createRow(linhaAtual++);
    		
	    		//=PROCV(MAIOR(A3:A6;1);A3:M6;2;FALSO)
	    		
	    		// Class
	    		if (i <= campeonato.getNumeroDestaques())
	    			escreveCell(row, CLASS_COL_CLASS, styleDestaqueCenter, (double) i);
	    		else
		    		escreveCell(row, CLASS_COL_CLASS, styleTabelaCenter, (double) i);

	    		// Time
	    		formula = "VLOOKUP(LARGE(" + intervaloMaior + "," + i + ")," + intervaloGrupo + "," + (CLASS_COL_TIME + 1) + ",FALSE)";
	    		escreveFormula(row, CLASS_COL_TIME, styleTabelaLeft, formula);

	    		// PG
	    		formula = "VLOOKUP(LARGE(" + intervaloMaior + "," + i + ")," + intervaloGrupo + "," + (CLASS_COL_PONTO_GANHO + 1) +  ",FALSE)";
	    		escreveFormula(row, CLASS_COL_PONTO_GANHO, styleTabelaCenter,formula);

	    		// PP
	    		formula = "VLOOKUP(LARGE(" + intervaloMaior + "," + i + ")," + intervaloGrupo + "," + (CLASS_COL_PONTO_PERDIDO + 1) +  ",FALSE)";
	    		escreveFormula(row, CLASS_COL_PONTO_PERDIDO, styleTabelaPontoPerdido,formula);

	    		// J
	    		formula = "VLOOKUP(LARGE(" + intervaloMaior + "," + i + ")," + intervaloGrupo + "," + (CLASS_COL_JOGO + 1) +  ",FALSE)";
	    		escreveFormula(row, CLASS_COL_JOGO, styleTabelaCenter,formula);
	    		
	    		// V
	    		formula = "VLOOKUP(LARGE(" + intervaloMaior + "," + i + ")," + intervaloGrupo + "," + (CLASS_COL_VITORIA + 1) +  ",FALSE)";
	    		escreveFormula(row, CLASS_COL_VITORIA, styleTabelaCenter,formula);

	    		// E
	    		formula = "VLOOKUP(LARGE(" + intervaloMaior + "," + i + ")," + intervaloGrupo + "," + (CLASS_COL_EMPATE + 1) +  ",FALSE)";
	    		escreveFormula(row, CLASS_COL_EMPATE, styleTabelaCenter,formula);

	    		// D
	    		formula = "VLOOKUP(LARGE(" + intervaloMaior + "," + i + ")," + intervaloGrupo + "," + (CLASS_COL_DERROTA + 1) +  ",FALSE)";
	    		escreveFormula(row, CLASS_COL_DERROTA, styleTabelaCenter,formula);

	    		// GP
	    		formula = "VLOOKUP(LARGE(" + intervaloMaior + "," + i + ")," + intervaloGrupo + "," + (CLASS_COL_GOL_PRO + 1) +  ",FALSE)";
	    		escreveFormula(row, CLASS_COL_GOL_PRO, styleTabelaCenter,formula);

	    		// GC
	    		formula = "VLOOKUP(LARGE(" + intervaloMaior + "," + i + ")," + intervaloGrupo + "," + (CLASS_COL_GOL_CONTRA + 1) +  ",FALSE)";
	    		escreveFormula(row, CLASS_COL_GOL_CONTRA, styleTabelaCenter,formula);

	    		// S
	    		formula = "VLOOKUP(LARGE(" + intervaloMaior + "," + i + ")," + intervaloGrupo + "," + (CLASS_COL_SALDO + 1) +  ",FALSE)";
	    		escreveFormula(row, CLASS_COL_SALDO, styleTabelaCenter,formula);

	    		// %
	    		formula = "VLOOKUP(LARGE(" + intervaloMaior + "," + i + ")," + intervaloGrupo + "," + (CLASS_COL_PORCENTAGEM + 1) +  ",FALSE)";
	    		escreveFormula(row, CLASS_COL_PORCENTAGEM, styleTabelaPorcentagem,formula);

	    		// Arbitragens
	    		if (campeonato.getIsArbitragem()) {
		    		formula = "VLOOKUP(LARGE(" + intervaloMaior + "," + i + ")," + intervaloGrupo + "," + (CLASS_COL_ARBITRAGEM + 1) +  ",FALSE)";
		    		escreveFormula(row, CLASS_COL_ARBITRAGEM, styleTabelaCenter,formula);
	    		}
	    

	    	}

	    }
	    
	    linhaAtual++;

	    // Total de jogos    
	    int totalPartidas = 0;
	    for (Grupo grupo : campeonato.getGrupos()) 
	    	totalPartidas += grupo.getTabelaJogos().size();
	    row = sheet.createRow(linhaAtual++);
	    escreveCell(row, COLUNA_ZERO, styleCorpoLeft, "Total de partidas:");
    	escreveCell(row, COLUNA_ZERO + 2, styleCorpoRight, (double) totalPartidas);
    	CellReference crTotalJogos = new CellReference(row.getRowNum(), COLUNA_ZERO + 2);

    	// jogos realizados
	    row = sheet.createRow(linhaAtual++);
    	escreveCell(row, COLUNA_ZERO, styleCorpoLeft, "Partidas realizadas:");
    	formula = "SUM(Auxiliar!" + retornaRange(campeonato, CLASS_COL_JOGO) + ")/2";
    	escreveFormula(row, COLUNA_ZERO + 2, styleCorpoRight, formula);
    	CellReference crJogosRealizados = new CellReference(row.getRowNum(), COLUNA_ZERO + 2);
 	    formula = crJogosRealizados.formatAsString() + "/" + crTotalJogos.formatAsString();
    	escreveFormula(row, COLUNA_ZERO + 3, styleCorpoPorcentagem, formula);

    	// jogos restantes
	    row = sheet.createRow(linhaAtual++);
    	escreveCell(row, COLUNA_ZERO, styleCorpoLeft, "Partidas restantes:");
    	formula = crTotalJogos.formatAsString() + "-" + crJogosRealizados.formatAsString();
    	escreveFormula(row, COLUNA_ZERO + 2, styleCorpoRight, formula);
    	CellReference crJogosRestantes = new CellReference(row.getRowNum(), COLUNA_ZERO + 2);
 	    formula = crJogosRestantes.formatAsString() + "/" + crTotalJogos.formatAsString();
    	escreveFormula(row, COLUNA_ZERO + 3, styleCorpoPorcentagem, formula);

    	// gols marcados
	    row = sheet.createRow(linhaAtual++);
    	escreveCell(row, COLUNA_ZERO, styleCorpoLeft, "Gols marcados:");
    	formula = "SUM(Auxiliar!" + retornaRange(campeonato, CLASS_COL_GOL_PRO) + ")";
    	escreveFormula(row, COLUNA_ZERO + 2, styleCorpoRight, formula);
    	CellReference crGolsMarcados = new CellReference(row.getRowNum(), COLUNA_ZERO + 2);

    	// Media de gols
	    row = sheet.createRow(linhaAtual++);
    	escreveCell(row, COLUNA_ZERO, styleCorpoLeft, "Média de gol:");
	    formula = "IF(" + crGolsMarcados.formatAsString() + ">0," + crGolsMarcados.formatAsString() + "/" + crJogosRealizados.formatAsString() + "," + AD + AD + ")";
    	escreveFormula(row, COLUNA_ZERO + 2, styleCorpoMedia, formula);

	
	}

	
	private static String retornaRange(Campeonato campeonato, int coluna) {
	    int linhaInicial = 0, linhaFinal = 0;
	    for (Grupo grupo : campeonato.getGrupos()) {
	    	int refLinhaInicial = grupo.getReferenciaInicialAux().getRow();
	    	if ((linhaInicial == 0) || (linhaInicial > refLinhaInicial)) 
	    		linhaInicial = refLinhaInicial;

	    	int refLinhaFinal = grupo.getReferenciaFinalAux().getRow();
	    	if ((linhaFinal == 0) || (linhaFinal < refLinhaFinal)) 
	    		linhaFinal = refLinhaFinal;
	    } 
	    return retornaColuna(coluna) + String.valueOf(linhaInicial + 1)+ ":" + retornaColuna(coluna) + String.valueOf(linhaFinal + 1);
	}

	private static void escreveCell(Row row, int cellNumber, XSSFCellStyle estilo, Object value) {
		Cell cell = row.createCell(cellNumber);
		cell.setCellStyle(estilo);

		if (value instanceof String) {
			cell.setCellValue((String) value);
		} else if(value instanceof Double){
			cell.setCellValue((Double) value);
		} else {
			cell.setCellValue(value.toString());
		}
	}
	
	private static void escreveFormula(Row row, int cellNumber, XSSFCellStyle estilo, String formula) {
		Cell cell = row.createCell(cellNumber);
		cell.setCellStyle(estilo);
	    cell.setCellFormula(formula);
	}
	
	public static void gravaPlanilha(XSSFWorkbook wb, Campeonato campeonato) {

	    //Write the workbook in file system
	    try {
	    	FileOutputStream out = new FileOutputStream(new File(campeonato.getNomeCampeonato() + ".xlsx"));
	    	wb.write(out);
	    	out.close();
	    	System.out.println("Arquivo " + campeonato.getNomeCampeonato() + ".xlsx gravado no disco");
	    } catch (Exception e) {
	    	e.printStackTrace();
	    }

	}

	public static void configuraStyles(XSSFWorkbook wb, Campeonato campeonato) {

		// Cores utilizadas
		XSSFColor corCabecalho = new XSSFColor(campeonato.getCorCabecalho(), null);
		XSSFColor corDestaque = new XSSFColor(campeonato.getCorDestaque(), null);
		
		//Fontes utilizadas
		XSSFFont fonteTitulo = wb.createFont();
		fonteTitulo.setFontHeightInPoints((short) 12);
		fonteTitulo.setFontName("Verdana");
		fonteTitulo.setBold(true);
	    
		XSSFFont fonteCabecalho = wb.createFont();
		fonteCabecalho.setFontHeightInPoints((short)8);
		fonteCabecalho.setFontName("Verdana");
		fonteCabecalho.setBold(true);

		XSSFFont fonteNormal = wb.createFont();
		fonteNormal.setFontHeightInPoints((short)8);
		fonteNormal.setFontName("Verdana");
		fonteNormal.setBold(false);

		XSSFFont fonteNormalRed = wb.createFont();
		fonteNormalRed.setFontHeightInPoints((short)8);
		fonteNormalRed.setFontName("Verdana");
		fonteNormalRed.setColor(IndexedColors.RED.index);
		fonteNormalRed.setBold(false);
		
		// Estilos
		styleTitulo = wb.createCellStyle();
		styleTitulo.setFont(fonteTitulo);

		styleFase = wb.createCellStyle();
		styleFase.setFont(fonteCabecalho);

		styleDestaqueCenter = wb.createCellStyle();
		styleDestaqueCenter.setFont(fonteNormal);
		styleDestaqueCenter.setFillForegroundColor(corDestaque);
		styleDestaqueCenter.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		styleDestaqueCenter.setBorderBottom(BorderStyle.THIN);
		styleDestaqueCenter.setBorderTop(BorderStyle.THIN);
		styleDestaqueCenter.setBorderLeft(BorderStyle.THIN);
		styleDestaqueCenter.setBorderRight(BorderStyle.THIN);
		styleDestaqueCenter.setAlignment(HorizontalAlignment.CENTER);
		
		styleCabecalho = wb.createCellStyle();
		styleCabecalho.setFont(fonteCabecalho);
		styleCabecalho.setFillForegroundColor(corCabecalho);
		styleCabecalho.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		styleCabecalho.setBorderBottom(BorderStyle.THIN);
		styleCabecalho.setBorderTop(BorderStyle.THIN);
		styleCabecalho.setBorderLeft(BorderStyle.THIN);
		styleCabecalho.setBorderRight(BorderStyle.THIN);

		styleCabecalhoCenter = wb.createCellStyle();
		styleCabecalhoCenter.cloneStyleFrom(styleCabecalho);
		styleCabecalhoCenter.setAlignment(HorizontalAlignment.CENTER);

		styleCabecalhoLeft = wb.createCellStyle();
		styleCabecalhoLeft.cloneStyleFrom(styleCabecalho);
		styleCabecalhoLeft.setAlignment(HorizontalAlignment.LEFT);

		styleCabecalhoRight = wb.createCellStyle();
		styleCabecalhoRight.cloneStyleFrom(styleCabecalho);
		styleCabecalhoRight.setAlignment(HorizontalAlignment.RIGHT);
		
		styleTabela = wb.createCellStyle();
		styleTabela.setFont(fonteNormal);
		styleTabela.setBorderBottom(BorderStyle.THIN);
		styleTabela.setBorderTop(BorderStyle.THIN);
		styleTabela.setBorderLeft(BorderStyle.THIN);
		styleTabela.setBorderRight(BorderStyle.THIN);

		styleTabelaCenter = wb.createCellStyle();
		styleTabelaCenter.cloneStyleFrom(styleTabela);
		styleTabelaCenter.setAlignment(HorizontalAlignment.CENTER);

		styleTabelaLeft = wb.createCellStyle();
		styleTabelaLeft.cloneStyleFrom(styleTabela);
		styleTabelaLeft.setAlignment(HorizontalAlignment.LEFT);

		styleTabelaRight = wb.createCellStyle();
		styleTabelaRight.cloneStyleFrom(styleTabela);
		styleTabelaRight.setAlignment(HorizontalAlignment.RIGHT);

		styleTabelaPontoPerdido = wb.createCellStyle();
		styleTabelaPontoPerdido.cloneStyleFrom(styleTabelaCenter);
		styleTabelaPontoPerdido.setFont(fonteNormalRed);
		
		styleTabelaPorcentagem = wb.createCellStyle();
		styleTabelaPorcentagem.cloneStyleFrom(styleTabelaRight);
		styleTabelaPorcentagem.setDataFormat(wb.createDataFormat().getFormat("0.00%"));
		styleTabelaPorcentagem.setFont(fonteNormal);

		
		styleCorpo = wb.createCellStyle();
		styleCorpo.cloneStyleFrom(styleTabela);
		styleCorpo.setBorderBottom(BorderStyle.NONE);
		styleCorpo.setBorderTop(BorderStyle.NONE);
		styleCorpo.setBorderLeft(BorderStyle.NONE);
		styleCorpo.setBorderRight(BorderStyle.NONE);

		styleCorpoPorcentagem = wb.createCellStyle();
		styleCorpoPorcentagem.cloneStyleFrom(styleCorpo);
		styleCorpoPorcentagem.setDataFormat(wb.createDataFormat().getFormat("0.00%"));
		styleCorpoPorcentagem.setFont(fonteNormal);
		
		styleCorpoCenter = wb.createCellStyle();
		styleCorpoCenter.cloneStyleFrom(styleCorpo);
		styleCorpoCenter.setAlignment(HorizontalAlignment.CENTER);

		styleCorpoLeft = wb.createCellStyle();
		styleCorpoLeft.cloneStyleFrom(styleCorpo);
		styleCorpoLeft.setAlignment(HorizontalAlignment.LEFT);

		styleCorpoRight = wb.createCellStyle();
		styleCorpoRight.cloneStyleFrom(styleCorpo);
		styleCorpoRight.setAlignment(HorizontalAlignment.RIGHT);

		styleCorpoMedia = wb.createCellStyle();
		styleCorpoMedia.cloneStyleFrom(styleCorpoRight);
		styleCorpoMedia.setDataFormat(wb.createDataFormat().getFormat("0.00"));
		styleCorpoMedia.setFont(fonteNormal);

		
	}	
	
	private static char retornaColuna(int numeroColuna) {
		String alfabeto = "ABCDEFGHIJKLMNOPQRSTUVXWYZ";
		return alfabeto.charAt(numeroColuna);
		
	}

}
