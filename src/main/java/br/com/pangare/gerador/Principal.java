package br.com.pangare.gerador;

import java.io.File;
import java.io.IOException;
import java.nio.charset.Charset;
import java.nio.file.Files;
import java.util.ArrayList;
import java.util.Collections;
import java.util.List;

import org.apache.commons.cli.CommandLine;
import org.apache.commons.cli.CommandLineParser;
import org.apache.commons.cli.DefaultParser;
import org.apache.commons.cli.HelpFormatter;
import org.apache.commons.cli.Options;
import org.apache.commons.cli.ParseException;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
* Geracao de tabelas em Excel para campeonatos de fubebol
*
* @author  Tarcizio Dinoa
* @version 1.0
* @since   2018-05-01 
*/

public class Principal {


	private final static String VERSAO_PROGRAMA = "GERADOR 1.0";

	private final static String DEFAULT_NOME_ARQUIVO = "lista.txt";
	private final static String DEFAULT_NOME_CAMPEONATO = "NOME DO CAMPEONATO";
	private final static byte[] DEFAULT_COR_CABECALHO = new byte[]{ (byte) 209, (byte) 196, (byte) 192};
	private final static byte[] DEFAULT_COR_DESTAQUE = new byte[]{ (byte) 247, (byte) 236, (byte) 227};
	private final static int DEFAULT_DESTAQUES = 2;

	private final static String LINHA_DUPLA = "===================";

	/**
	 * @param args
	 * @throws ParseException 
	 * @throws IOException 
	 */
	public static void main(String[] args) throws ParseException, IOException  {

		Campeonato campeonato = trata_parametros(args);
		
		System.out.println(VERSAO_PROGRAMA);
		System.out.println(LINHA_DUPLA);

		System.out.println("Nome do torneio: "+ campeonato.getNomeCampeonato());
		System.out.println("Ida e volta: " + campeonato.getIsDoisTurnos());
		System.out.println("Arbitragem: " + campeonato.getIsArbitragem());
		System.out.println("Destaques: " + campeonato.getNumeroDestaques());
		System.out.println("Valores aleatorios: " + campeonato.getIsValoresAleatorios());
		System.out.println("Lendo lista de grupos e clubes a partir de " + campeonato.getFileName());
		
		// Processa grupos
		List<Grupo> grupos = new ArrayList<Grupo>();

		// Le grupos do arquivo
		List<String> linhaGrupos = Files.readAllLines(new File(campeonato.getFileName()).toPath(), Charset.defaultCharset());

		int index = 0;

		for (String linha : linhaGrupos) {

			if (linha.trim().startsWith("#") || linha.length() == 0) // trata comentarios e linhas em branco
				continue;

			List<String> timesGrupo = new ArrayList<String>();
			String[] times = linha.split(",");
			for (String time : times) {
				timesGrupo.add(time.trim());
			}
				
			Grupo grupo = new Grupo();
			grupo.setNomeGrupo(retornaNomeGrupo(index++)); 

			Collections.shuffle(timesGrupo); // embaralha os grupos
			grupo.setTimes(timesGrupo);
			grupos.add(grupo);
				
		}

		campeonato.setGrupos(grupos);
			
		System.out.println("Numero de grupos: " + campeonato.getGrupos().size());
		
		for (Grupo grupo : campeonato.getGrupos()) {
			System.out.println("Gerando partidas para o " + grupo.getNomeGrupo() + "...");
		    grupo.setTabelaJogos(geradorPartidas(grupo, campeonato.getIsDoisTurnos()));
		}

//		listaJogos(campeonato);

		XSSFWorkbook wb = new XSSFWorkbook();

		System.out.println("Criando aba de jogos...");
		GravaXLS.criaSheetJogos(wb, campeonato);

		System.out.println("Criando aba auxiliar...");
		GravaXLS.criaSheetAuxiliar(wb, campeonato);
		// Oculta aba auxiliar
		wb.setSheetHidden(wb.getSheetIndex("Auxiliar"), true);
	
		System.out.println("Criando aba de classificacao...");
		GravaXLS.criaSheetClassificacao(wb, campeonato);

		System.out.println("Gravando planilha...");
		GravaXLS.gravaPlanilha(wb, campeonato);

		System.out.println(LINHA_DUPLA);
		System.out.println("Fim de processamento.");
		
	}


	private static Campeonato trata_parametros(String[] args) throws ParseException {

		Options options = getOptions(); 
		CommandLineParser parser = new DefaultParser();
		CommandLine cmd = parser.parse(options, args);

		Campeonato campeonato = new Campeonato();
		
		campeonato.setNomeCampeonato(DEFAULT_NOME_CAMPEONATO); // default
		campeonato.setIsArbitragem(true); // default
		campeonato.setIsDoisTurnos(false); // default
		campeonato.setIsValoresAleatorios(false); // default
		campeonato.setCorCabecalho(DEFAULT_COR_CABECALHO); // default
		campeonato.setCorDestaque(DEFAULT_COR_DESTAQUE); // default
		campeonato.setNumeroDestaques(DEFAULT_DESTAQUES); // default
		campeonato.setFileName(DEFAULT_NOME_ARQUIVO);
		
		// Trata parametros		
		if (cmd.hasOption("h")) {
			HelpFormatter formatter = new HelpFormatter();
			formatter.printHelp("gerador", options );
			System.exit(1);
		}
		if (cmd.hasOption("f")) 
			campeonato.setFileName(cmd.getOptionValue("f"));
		if (cmd.hasOption("nome")) 
			campeonato.setNomeCampeonato(cmd.getOptionValue("nome"));
		if (cmd.hasOption("nd")) 
			campeonato.setNumeroDestaques(Integer.parseInt(cmd.getOptionValue("nd")));
		if (cmd.hasOption("vv")) 
			campeonato.setIsDoisTurnos(true);
		if (cmd.hasOption("va")) 
			campeonato.setIsValoresAleatorios(true);
		if (cmd.hasOption("sa")) 
			campeonato.setIsArbitragem(false);
		if (cmd.hasOption("cor_cabecalho")) {
			String[] cores = cmd.getOptionValue("cor_cabecalho").split(",");
			if (cores.length == 3) {
				try {
					byte[] coresCabecalho = new byte[]{(byte) Integer.parseInt(cores[0].trim()), (byte) Integer.parseInt(cores[1].trim()), (byte) Integer.parseInt(cores[2].trim())};
					campeonato.setCorCabecalho(coresCabecalho);
				} catch (NumberFormatException e) {
					// TODO: handle exception
				}
			}
		}
		if (cmd.hasOption("cor_destaque")) {
			String[] cores = cmd.getOptionValue("cor_destaque").split(",");
			if (cores.length == 3) {
				try {
					byte[] coresDestaque = new byte[]{(byte) Integer.parseInt(cores[0].trim()), (byte) Integer.parseInt(cores[1].trim()), (byte) Integer.parseInt(cores[2].trim())};
					campeonato.setCorDestaque(coresDestaque);
				} catch (NumberFormatException e) {
					// TODO: handle exception
				}
			}
		}
		
		
		return campeonato;
		
	}


	/**
	 * Retorna uma lista de partidas a partir da lista de clubes informada.
	 * @param List<String> clubes 
	 * @param Boolean isDoisTurnos
	 * @return ArrayList<Jogo>
	 */
	private static ArrayList<Jogo> geradorPartidas(Grupo grupo, boolean ida_volta) {
		
		ArrayList<Jogo> tabelaJogos = new ArrayList<Jogo>();

		if (grupo.getTimes().size() % 2 == 1) {
	        grupo.getTimes().add(0, "");
	    }

		int numClubes = grupo.getTimes().size();
	    int numJogosPorRodada = numClubes / 2;
	    int numJogo = 1;
	    int rodadaAtual = 0;
	    
	    for (int i = 0; i < numClubes - 1; i++) {

	        for (int j = 0; j < numJogosPorRodada; j++) {

	        	// Clube esta de fora nessa rodada?              
	        	if (grupo.getTimes().get(j).isEmpty()) {
	        		continue;
	        	}

	        	rodadaAtual = i + 1;
	        	
				Jogo jogo = new Jogo();
				jogo.setRodada(rodadaAtual);
				jogo.setNumJogo(numJogo++);
	        	
	        	// Teste para ajustar o mando de campo
	            if (j % 2 == 1 || i % 2 == 1 && j == 0) {
					jogo.setTimeMandante(grupo.getTimes().get(numClubes - j - 1));
					jogo.setTimeVisitante(grupo.getTimes().get(j));
	            } else {
					jogo.setTimeMandante(grupo.getTimes().get(j));
					jogo.setTimeVisitante(grupo.getTimes().get(numClubes - j - 1));
	            }
	            
				tabelaJogos.add(jogo);
				
	        }
	        // Gira os clubes no sentido hor�rio, mantendo o primeiro no lugar
	        grupo.getTimes().add(1,grupo.getTimes().remove(grupo.getTimes().size()-1));
	    }

		if (ida_volta) {
			
			ArrayList<Jogo> tabelaJogosClone = new ArrayList<Jogo>(tabelaJogos);
			Collections.copy(tabelaJogosClone, tabelaJogos);
			
			for (Jogo jogo : tabelaJogosClone) {
				Jogo novoJogo = new Jogo();
				novoJogo.setTimeMandante(jogo.getTimeVisitante());
				novoJogo.setTimeVisitante(jogo.getTimeMandante());
				novoJogo.setRodada(rodadaAtual + jogo.getRodada());
				novoJogo.setNumJogo(numJogo++);
				tabelaJogos.add(novoJogo);
			}
			
		}

	    if (grupo.getTimes().get(0) == "") {
	    	grupo.getTimes().remove(0);
			grupo.setTimes(grupo.getTimes());
	    }

		return tabelaJogos;
		
	}

	
	private static void listaJogos(Campeonato campeonato) {
		List<Grupo> grupos = campeonato.getGrupos();
		for (Grupo grupo : grupos) {
			System.out.println("- " + grupo.getNomeGrupo() + ":");
			for (Jogo jogo : grupo.getTabelaJogos()) {
				System.out.println(jogo.getRodada() + " - " + jogo.getNumJogo() + " - " + jogo.getTimeMandante() + " x " + jogo.getTimeVisitante());
			}
		}
	}


	private static Options getOptions() {
		  Options options = new Options();
		  options.addOption("h", false, "Apresenta esta tela");
		  options.addOption("vv",false, "Jogos de ida e volta");
		  options.addOption("va",false, "Preenche os jogos com valores aleatorios (0-4)");
		  options.addOption("f", true, "Arquivo com lista de clubes [default: " + DEFAULT_NOME_ARQUIVO + "]");
		  options.addOption("nome", true, "Nome do campeonato [default: " + DEFAULT_NOME_CAMPEONATO + "]");
		  options.addOption("sa", false, "Sem coluna de arbitragem");
		  options.addOption("cor_cabecalho", true, "Cores RGB para cabecalho [default: 209,196,192]");
		  options.addOption("cor_destaque", true, "Cores RGB para destaque [default: 247,236,227]");
		  options.addOption("nd", true, "Numero de clubes em destaque na classificação [default: 2]");
		  return options;
	}


	private static String retornaNomeGrupo(int indiceGrupo) {
		String alfabeto = "ABCDEFGHIJKLMNOPQRSTUVXWYZ";
		return "Grupo " + alfabeto.charAt(indiceGrupo);
	}	

}
