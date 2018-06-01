package br.com.pangare.gerador;

import java.util.ArrayList;
import java.util.List;

import org.apache.poi.ss.util.CellReference;

public class Grupo {

	private String nomeGrupo;
	private List<String> times;
	private ArrayList<Jogo> tabelaJogos;
	private CellReference referenciaInicialAux;
	private CellReference referenciaUltimaLinhaAux;
	private CellReference referenciaFinalAux;

	/**
	 * @param nomeGrupo the nomeGrupo to set
	 */
	public void setNomeGrupo(String nomeGrupo) {
		this.nomeGrupo = nomeGrupo;
	}
	/**
	 * @return the nomeGrupo
	 */
	public String getNomeGrupo() {
		return nomeGrupo;
	}
	/**
	 * @param times the times to set
	 */
	public void setTimes(List<String> times) {
		this.times = times;
	}
	/**
	 * @return the times
	 */
	public List<String> getTimes() {
		return times;
	}
	/**
	 * @param tabelaJogos the tabelaJogos to set
	 */
	public void setTabelaJogos(ArrayList<Jogo> tabelaJogos) {
		this.tabelaJogos = tabelaJogos;
	}
	/**
	 * @return the tabelaJogos
	 */
	public ArrayList<Jogo> getTabelaJogos() {
		return tabelaJogos;
	}
	/**
	 * @return the referenciaInicialAux
	 */
	public CellReference getReferenciaInicialAux() {
		return referenciaInicialAux;
	}
	/**
	 * @param referenciaInicialAux the referenciaInicialAux to set
	 */
	public void setReferenciaInicialAux(CellReference referenciaInicialAux) {
		this.referenciaInicialAux = referenciaInicialAux;
	}
	/**
	 * @return the referenciaFinalAux
	 */
	public CellReference getReferenciaFinalAux() {
		return referenciaFinalAux;
	}
	/**
	 * @param referenciaFinalAux the referenciaFinalAux to set
	 */
	public void setReferenciaFinalAux(CellReference referenciaFinalAux) {
		this.referenciaFinalAux = referenciaFinalAux;
	}
	/**
	 * @return the referenciaUltimaLinhaAux
	 */
	public CellReference getReferenciaUltimaLinhaAux() {
		return referenciaUltimaLinhaAux;
	}
	/**
	 * @param referenciaUltimaLinhaAux the referenciaUltimaLinhaAux to set
	 */
	public void setReferenciaUltimaLinhaAux(CellReference referenciaUltimaLinhaAux) {
		this.referenciaUltimaLinhaAux = referenciaUltimaLinhaAux;
	}




}
