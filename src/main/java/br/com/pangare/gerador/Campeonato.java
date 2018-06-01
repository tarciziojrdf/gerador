package br.com.pangare.gerador;

import java.util.List;

public class Campeonato {

	private String nomeCampeonato;
	private String fileName;
	private Boolean isDoisTurnos;
	private Boolean isArbitragem;
	private Boolean isValoresAleatorios;
	private List<Grupo> grupos;
	private int numeroGrupos;
	private int numeroDestaques;
	private byte[] corCabecalho;
	private byte[] corDestaque;
	
	
	/**
	 * @param nomeCampeonato the nomeCampeonato to set
	 */
	public void setNomeCampeonato(String nomeCampeonato) {
		this.nomeCampeonato = nomeCampeonato;
	}
	/**
	 * @return the nomeCampeonato
	 */
	public String getNomeCampeonato() {
		return nomeCampeonato;
	}
	/**
	 * @param isDoisTurnos the isDoisTurnos to set
	 */
	public void setIsDoisTurnos(boolean isDoisTurnos) {
		this.isDoisTurnos = isDoisTurnos;
	}
	/**
	 * @return the isDoisTurnos
	 */
	public boolean getIsDoisTurnos() {
		return isDoisTurnos.booleanValue();
	}
	/**
	 * @param isArbitragem the isArbitragem to set
	 */
	public void setIsArbitragem(boolean isArbitragem) {
		this.isArbitragem = isArbitragem;
	}
	/**
	 * @return the isArbitragem
	 */
	public boolean getIsArbitragem() {
		return isArbitragem.booleanValue();
	}
	/**
	 * @param isValoresAleatorios the isValoresAleatorios to set
	 */
	public void setIsValoresAleatorios(Boolean isValoresAleatorios) {
		this.isValoresAleatorios = isValoresAleatorios;
	}
	/**
	 * @return the isValoresAleatorios
	 */
	public Boolean getIsValoresAleatorios() {
		return isValoresAleatorios;
	}
	/**
	 * @param numeroGrupos the numeroGrupos to set
	 */
	public void setNumeroGrupos(int numeroGrupos) {
		this.numeroGrupos = numeroGrupos;
	}
	/**
	 * @return the numeroGrupos
	 */
	public int getNumeroGrupos() {
		return numeroGrupos;
	}
	/**
	 * @param grupos the grupos to set
	 */
	public void setGrupos(List<Grupo> grupos) {
		this.grupos = grupos;
	}
	/**
	 * @return the grupos
	 */
	public List<Grupo> getGrupos() {
		return grupos;
	}
	/**
	 * @return the corCabecalho
	 */
	public byte[] getCorCabecalho() {
		return corCabecalho;
	}
	/**
	 * @param corCabecalho the corCabecalho to set
	 */
	public void setCorCabecalho(byte[] corCabecalho) {
		this.corCabecalho = corCabecalho;
	}
	/**
	 * @return the numeroDestaques
	 */
	public int getNumeroDestaques() {
		return numeroDestaques;
	}
	/**
	 * @param numeroDestaques the numeroDestaques to set
	 */
	public void setNumeroDestaques(int numeroDestaques) {
		this.numeroDestaques = numeroDestaques;
	}
	/**
	 * @return the corDestaque
	 */
	public byte[] getCorDestaque() {
		return corDestaque;
	}
	/**
	 * @param corDestaque the corDestaque to set
	 */
	public void setCorDestaque(byte[] corDestaque) {
		this.corDestaque = corDestaque;
	}
	/**
	 * @return the fileName
	 */
	public String getFileName() {
		return fileName;
	}
	/**
	 * @param fileName the fileName to set
	 */
	public void setFileName(String fileName) {
		this.fileName = fileName;
	}


}
