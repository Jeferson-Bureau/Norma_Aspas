/**
 * Sistema de Formatação APA 7
 * Tipos e Interfaces TypeScript
 */

// ============================================================================
// TIPOS BASE
// ============================================================================

export type TipoReferencia =
  | 'artigo'
  | 'livro'
  | 'capitulo'
  | 'website'
  | 'tese'
  | 'dissertacao'
  | 'jornal'
  | 'video'
  | 'podcast'
  | 'conferencia'
  | 'relatorio'
  | 'lei'
  | 'midia-social';

export type TipoCitacao = 'narrativa' | 'parentetica' | 'direta-curta' | 'direta-longa';

// ============================================================================
// INTERFACES DE AUTOR
// ============================================================================

export interface Autor {
  nome: string;
  sobrenome: string;
  inicialMeio?: string;
}

export interface AutorCorporativo {
  nomeCompleto: string;
  sigla?: string;
}

// ============================================================================
// INTERFACE PRINCIPAL DE REFERÊNCIA
// ============================================================================

export interface ReferenciaBase {
  tipo: TipoReferencia;
  autores?: Autor[];
  autorCorporativo?: AutorCorporativo;
  ano: string;
  titulo: string;
  doi?: string;
  url?: string;
  dataAcesso?: Date;
  idioma?: 'pt' | 'en' | 'es';
}

// ============================================================================
// INTERFACES ESPECÍFICAS POR TIPO
// ============================================================================

export interface ReferenciaArtigo extends ReferenciaBase {
  tipo: 'artigo';
  periodico: string;
  volume: string;
  numero?: string;
  paginas: string;
  issn?: string;
}

export interface ReferenciaLivro extends ReferenciaBase {
  tipo: 'livro';
  editora: string;
  edicao?: string;
  local?: string;
  isbn?: string;
}

export interface ReferenciaCapitulo extends ReferenciaBase {
  tipo: 'capitulo';
  editores: Autor[];
  tituloLivro: string;
  editora: string;
  paginas: string;
  edicao?: string;
  isbn?: string;
}

export interface ReferenciaWebsite extends ReferenciaBase {
  tipo: 'website';
  nomeSite: string;
  dataPublicacao?: string;
}

export interface ReferenciaTese extends ReferenciaBase {
  tipo: 'tese' | 'dissertacao';
  instituicao: string;
  tipoTrabalho: 'Tese de doutorado' | 'Dissertação de mestrado';
  repositorio?: string;
}

export interface ReferenciaJornal extends ReferenciaBase {
  tipo: 'jornal';
  nomeJornal: string;
  dataPublicacao: string; // formato: YYYY, Mês DD
  paginas?: string;
}

export interface ReferenciaVideo extends ReferenciaBase {
  tipo: 'video';
  plataforma: 'YouTube' | 'Vimeo' | 'Outro';
  canal?: string;
  duracao?: string;
  dataPublicacao: string;
}

export interface ReferenciaPodcast extends ReferenciaBase {
  tipo: 'podcast';
  apresentador: string;
  numeroEpisodio?: string;
  nomePodcast: string;
  produtora?: string;
  dataPublicacao: string;
}

export interface ReferenciaConferencia extends ReferenciaBase {
  tipo: 'conferencia';
  nomeConferencia: string;
  local: string;
  dataInicio: string;
  dataFim?: string;
  tipoApresentacao: 'Apresentação de artigo' | 'Pôster' | 'Palestra' | 'Simpósio';
}

export interface ReferenciaRelatorio extends ReferenciaBase {
  tipo: 'relatorio';
  numeroRelatorio?: string;
  instituicao: string;
  editora?: string;
}

export interface ReferenciaLei extends ReferenciaBase {
  tipo: 'lei';
  numeroLei: string;
  secao?: string;
}

export interface ReferenciaMidiaSocial extends ReferenciaBase {
  tipo: 'midia-social';
  plataforma: 'Twitter' | 'Facebook' | 'Instagram' | 'LinkedIn';
  username: string;
  conteudo: string; // Limitado a 20 palavras
  tipoPost: 'Tweet' | 'Post' | 'Story';
}

// ============================================================================
// UNION TYPE DE TODAS AS REFERÊNCIAS
// ============================================================================

export type Referencia =
  | ReferenciaArtigo
  | ReferenciaLivro
  | ReferenciaCapitulo
  | ReferenciaWebsite
  | ReferenciaTese
  | ReferenciaJornal
  | ReferenciaVideo
  | ReferenciaPodcast
  | ReferenciaConferencia
  | ReferenciaRelatorio
  | ReferenciaLei
  | ReferenciaMidiaSocial;

// ============================================================================
// INTERFACE DE CITAÇÃO
// ============================================================================

export interface CitacaoInput {
  referencia: Referencia;
  tipo: TipoCitacao;
  pagina?: string | number;
  paginas?: string; // Ex: "45-47"
  paragrafo?: number;
  textoCitado?: string; // Para citações diretas
  sufixo?: string; // Para múltiplas obras do mesmo autor no mesmo ano (a, b, c)
}

// ============================================================================
// INTERFACE DE OUTPUT
// ============================================================================

export interface ReferenciaFormatada {
  referenciaCompleta: string;
  citacaoNarrativa: string;
  citacaoParentetica: string;
  tipoDetectado: TipoReferencia;
  validacao: ValidacaoReferencia;
  metadados: MetadadosReferencia;
}

export interface ValidacaoReferencia {
  valida: boolean;
  completa: boolean;
  camposObrigatoriosFaltantes: string[];
  camposOpcionaisFaltantes: string[];
  avisos: string[];
  erros: string[];
}

export interface MetadadosReferencia {
  temDOI: boolean;
  temURL: boolean;
  temISBN?: boolean;
  temISSN?: boolean;
  quantidadeAutores: number;
  contemAutorCorporativo: boolean;
  edicao?: string;
  idioma: 'pt' | 'en' | 'es';
}

// ============================================================================
// INTERFACE DE CONFIGURAÇÃO
// ============================================================================

export interface ConfiguracaoAPA {
  idiomaPadrao: 'pt' | 'en' | 'es';
  formatoData: 'completo' | 'ano-mes' | 'ano';
  incluirDOI: boolean;
  incluirURL: boolean;
  incluirDataAcesso: boolean;
  validacaoEstrita: boolean;
  usarEtAl: {
    apartirDeQuantosAutores: number; // Padrão: 3
    primeiraVez: boolean; // Se true, usa et al. desde primeira citação para 3+ autores
  };
}

// ============================================================================
// INTERFACE PARA MÚLTIPLAS CITAÇÕES
// ============================================================================

export interface CitacaoMultipla {
  referencias: Array<{
    referencia: Referencia;
    sufixo?: string;
    pagina?: string;
  }>;
  ordenarAlfabeticamente: boolean;
}

// ============================================================================
// CONSTANTES E ENUMS
// ============================================================================

export const MESES_PT = {
  1: 'janeiro',
  2: 'fevereiro',
  3: 'março',
  4: 'abril',
  5: 'maio',
  6: 'junho',
  7: 'julho',
  8: 'agosto',
  9: 'setembro',
  10: 'outubro',
  11: 'novembro',
  12: 'dezembro',
} as const;

export const MESES_EN = {
  1: 'January',
  2: 'February',
  3: 'March',
  4: 'April',
  5: 'May',
  6: 'June',
  7: 'July',
  8: 'August',
  9: 'September',
  10: 'October',
  11: 'November',
  12: 'December',
} as const;

export const ABREVIACOES = {
  ed: 'ed.',
  eds: 'eds.',
  ed_revisada: 'ed. rev.',
  pagina: 'p.',
  paginas: 'pp.',
  paragrafo: 'para.',
  volume: 'vol.',
  numero: 'n.',
  sem_data_pt: 's.d.',
  sem_data_en: 'n.d.',
  et_al: 'et al.',
  traducao: 'Trad.',
  organizador: 'Org.',
  coordenador: 'Coord.',
} as const;

// ============================================================================
// INTERFACE DE ERRO CUSTOMIZADO
// ============================================================================

export class ErroFormatacaoAPA extends Error {
  constructor(
    public tipo: 'campo_obrigatorio' | 'formato_invalido' | 'tipo_invalido',
    public campo: string,
    message: string
  ) {
    super(message);
    this.name = 'ErroFormatacaoAPA';
  }
}

// ============================================================================
// TYPE GUARDS
// ============================================================================

export function isReferenciaArtigo(ref: Referencia): ref is ReferenciaArtigo {
  return ref.tipo === 'artigo';
}

export function isReferenciaLivro(ref: Referencia): ref is ReferenciaLivro {
  return ref.tipo === 'livro';
}

export function isReferenciaCapitulo(ref: Referencia): ref is ReferenciaCapitulo {
  return ref.tipo === 'capitulo';
}

export function isReferenciaWebsite(ref: Referencia): ref is ReferenciaWebsite {
  return ref.tipo === 'website';
}

export function isReferenciaTese(ref: Referencia): ref is ReferenciaTese {
  return ref.tipo === 'tese' || ref.tipo === 'dissertacao';
}

export function isReferenciaVideo(ref: Referencia): ref is ReferenciaVideo {
  return ref.tipo === 'video';
}

export function isReferenciaPodcast(ref: Referencia): ref is ReferenciaPodcast {
  return ref.tipo === 'podcast';
}
