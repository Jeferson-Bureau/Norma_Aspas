/**
 * Sistema de Formatação APA 7 - Classe Principal
 * Responsável pela formatação de referências bibliográficas
 */

import {
    Referencia,
    ReferenciaFormatada,
    ConfiguracaoAPA,
    ValidacaoReferencia,
    MetadadosReferencia,
    Autor,
    ErroFormatacaoAPA,
    isReferenciaArtigo,
    isReferenciaLivro,
    isReferenciaCapitulo,
    isReferenciaWebsite,
    isReferenciaTese,
    isReferenciaVideo,
    isReferenciaPodcast,
    ABREVIACOES,
} from './types';

export class FormatadorAPA7 {
    private config: ConfiguracaoAPA;

    constructor(config?: Partial<ConfiguracaoAPA>) {
        this.config = {
            idiomaPadrao: config?.idiomaPadrao || 'pt',
            formatoData: config?.formatoData || 'completo',
            incluirDOI: config?.incluirDOI ?? true,
            incluirURL: config?.incluirURL ?? true,
            incluirDataAcesso: config?.incluirDataAcesso ?? false,
            validacaoEstrita: config?.validacaoEstrita ?? true,
            usarEtAl: {
                apartirDeQuantosAutores: config?.usarEtAl?.apartirDeQuantosAutores || 3,
                primeiraVez: config?.usarEtAl?.primeiraVez ?? false,
            },
        };
    }

    // ==========================================================================
    // MÉTODO PRINCIPAL DE FORMATAÇÃO
    // ==========================================================================

    public formatar(referencia: Referencia): ReferenciaFormatada {
        // Validar antes de formatar
        const validacao = this.validar(referencia);

        if (!validacao.valida && this.config.validacaoEstrita) {
            throw new ErroFormatacaoAPA(
                'campo_obrigatorio',
                validacao.camposObrigatoriosFaltantes[0] || 'campo desconhecido',
                `Campo obrigatório ausente: ${validacao.camposObrigatoriosFaltantes[0]}`
            );
        }

        // Formatar de acordo com o tipo
        let referenciaCompleta: string;

        if (isReferenciaArtigo(referencia)) {
            referenciaCompleta = this.formatarArtigo(referencia);
        } else if (isReferenciaLivro(referencia)) {
            referenciaCompleta = this.formatarLivro(referencia);
        } else if (isReferenciaCapitulo(referencia)) {
            referenciaCompleta = this.formatarCapitulo(referencia);
        } else if (isReferenciaWebsite(referencia)) {
            referenciaCompleta = this.formatarWebsite(referencia);
        } else if (isReferenciaTese(referencia)) {
            referenciaCompleta = this.formatarTese(referencia);
        } else if (isReferenciaVideo(referencia)) {
            referenciaCompleta = this.formatarVideo(referencia);
        } else if (isReferenciaPodcast(referencia)) {
            referenciaCompleta = this.formatarPodcast(referencia);
        } else {
            throw new ErroFormatacaoAPA(
                'tipo_invalido',
                'tipo',
                `Tipo de referência não suportado: ${referencia.tipo}`
            );
        }

        // Gerar citações
        const citacaoNarrativa = this.gerarCitacaoNarrativa(referencia);
        const citacaoParentetica = this.gerarCitacaoParentetica(referencia);

        // Gerar metadados
        const metadados = this.gerarMetadados(referencia);

        return {
            referenciaCompleta,
            citacaoNarrativa,
            citacaoParentetica,
            tipoDetectado: referencia.tipo,
            validacao,
            metadados,
        };
    }

    // ==========================================================================
    // FORMATADORES ESPECÍFICOS POR TIPO
    // ==========================================================================

    private formatarArtigo(ref: Referencia & { tipo: 'artigo' }): string {
        const partes: string[] = [];

        // Autores
        partes.push(this.formatarAutores(ref.autores || [], ref.autorCorporativo));

        // Ano
        partes.push(`(${ref.ano}).`);

        // Título do artigo (sentence case, sem itálico)
        partes.push(this.aplicarSentenceCase(ref.titulo) + '.');

        // Nome do periódico (itálico, title case para inglês)
        const periodico = ref.idioma === 'en'
            ? this.aplicarTitleCase(ref.periodico)
            : ref.periodico;
        partes.push(`*${periodico}*,`);

        // Volume (itálico)
        const volumeNumero = ref.numero
            ? `*${ref.volume}*(${ref.numero}),`
            : `*${ref.volume}*,`;
        partes.push(volumeNumero);

        // Páginas
        partes.push(`${ref.paginas}.`);

        // DOI ou URL
        if (this.config.incluirDOI && ref.doi) {
            partes.push(`https://doi.org/${ref.doi}`);
        } else if (this.config.incluirURL && ref.url) {
            partes.push(ref.url);
        }

        return partes.join(' ');
    }

    private formatarLivro(ref: Referencia & { tipo: 'livro' }): string {
        const partes: string[] = [];

        // Autores
        partes.push(this.formatarAutores(ref.autores || [], ref.autorCorporativo));

        // Ano
        partes.push(`(${ref.ano}).`);

        // Título (itálico, sentence case)
        const titulo = this.aplicarSentenceCase(ref.titulo);

        // Edição (se não for a primeira)
        if (ref.edicao && ref.edicao !== '1') {
            partes.push(`*${titulo}* (${ref.edicao}ª ${ABREVIACOES.ed}).`);
        } else {
            partes.push(`*${titulo}*.`);
        }

        // Editora
        partes.push(`${ref.editora}.`);

        // DOI ou URL
        if (this.config.incluirDOI && ref.doi) {
            partes.push(`https://doi.org/${ref.doi}`);
        } else if (this.config.incluirURL && ref.url) {
            partes.push(ref.url);
        }

        return partes.join(' ');
    }

    private formatarCapitulo(ref: Referencia & { tipo: 'capitulo' }): string {
        const partes: string[] = [];

        // Autores do capítulo
        partes.push(this.formatarAutores(ref.autores || [], ref.autorCorporativo));

        // Ano
        partes.push(`(${ref.ano}).`);

        // Título do capítulo (sentence case, sem itálico)
        partes.push(this.aplicarSentenceCase(ref.titulo) + '.');

        // In + Editores
        const editoresFormatados = this.formatarEditores(ref.editores);
        partes.push(`In ${editoresFormatados}`);

        // Título do livro (itálico, sentence case)
        const tituloLivro = this.aplicarSentenceCase(ref.tituloLivro);

        // Com edição se houver
        if (ref.edicao && ref.edicao !== '1') {
            partes.push(`*${tituloLivro}* (${ref.edicao}ª ${ABREVIACOES.ed},`);
        } else {
            partes.push(`*${tituloLivro}*`);
        }

        // Páginas
        partes.push(`(${ABREVIACOES.paginas} ${ref.paginas}).`);

        // Editora
        partes.push(`${ref.editora}.`);

        // DOI ou URL
        if (this.config.incluirDOI && ref.doi) {
            partes.push(`https://doi.org/${ref.doi}`);
        } else if (this.config.incluirURL && ref.url) {
            partes.push(ref.url);
        }

        return partes.join(' ');
    }

    private formatarWebsite(ref: Referencia & { tipo: 'website' }): string {
        const partes: string[] = [];

        // Autor (se diferente do nome do site)
        if (ref.autores && ref.autores.length > 0) {
            partes.push(this.formatarAutores(ref.autores, undefined));
        } else if (ref.autorCorporativo && ref.autorCorporativo.nomeCompleto !== ref.nomeSite) {
            partes.push(`${ref.autorCorporativo.nomeCompleto}.`);
        }

        // Ano
        const ano = ref.ano === 's.d.' || ref.ano === 'n.d.' ? ref.ano : ref.ano;
        const dataCompleta = ref.dataPublicacao || '';

        if (dataCompleta) {
            partes.push(`(${dataCompleta}).`);
        } else {
            partes.push(`(${ano}).`);
        }

        // Título (itálico, sentence case)
        partes.push(`*${this.aplicarSentenceCase(ref.titulo)}*.`);

        // Nome do site (se diferente do autor)
        if (!ref.autores || ref.autores.length === 0) {
            if (!ref.autorCorporativo || ref.autorCorporativo.nomeCompleto === ref.nomeSite) {
                // Não repetir o nome do site
            } else {
                partes.push(`${ref.nomeSite}.`);
            }
        } else {
            partes.push(`${ref.nomeSite}.`);
        }

        // URL
        if (ref.url) {
            partes.push(ref.url);
        }

        return partes.join(' ');
    }

    private formatarTese(ref: Referencia & { tipo: 'tese' | 'dissertacao' }): string {
        const partes: string[] = [];

        // Autores
        partes.push(this.formatarAutores(ref.autores || [], ref.autorCorporativo));

        // Ano
        partes.push(`(${ref.ano}).`);

        // Título (itálico, sentence case)
        partes.push(`*${this.aplicarSentenceCase(ref.titulo)}*`);

        // Tipo de trabalho e instituição
        partes.push(`[${ref.tipoTrabalho}, ${ref.instituicao}].`);

        // Repositório (se houver)
        if (ref.repositorio) {
            partes.push(`${ref.repositorio}.`);
        }

        // DOI ou URL
        if (this.config.incluirDOI && ref.doi) {
            partes.push(`https://doi.org/${ref.doi}`);
        } else if (this.config.incluirURL && ref.url) {
            partes.push(ref.url);
        }

        return partes.join(' ');
    }

    private formatarVideo(ref: Referencia & { tipo: 'video' }): string {
        const partes: string[] = [];

        // Autor/Canal
        if (ref.autores && ref.autores.length > 0) {
            partes.push(this.formatarAutores(ref.autores, undefined));
        } else if (ref.canal) {
            partes.push(`${ref.canal}.`);
        }

        // Data de publicação
        partes.push(`(${ref.dataPublicacao}).`);

        // Título (itálico, sentence case)
        partes.push(`*${this.aplicarSentenceCase(ref.titulo)}*`);

        // Descritor
        partes.push('[Vídeo].');

        // Plataforma
        partes.push(`${ref.plataforma}.`);

        // URL
        if (ref.url) {
            partes.push(ref.url);
        }

        return partes.join(' ');
    }

    private formatarPodcast(ref: Referencia & { tipo: 'podcast' }): string {
        const partes: string[] = [];

        // Apresentador
        partes.push(`${ref.apresentador} (Host).`);

        // Data de publicação
        partes.push(`(${ref.dataPublicacao}).`);

        // Título do episódio
        partes.push(this.aplicarSentenceCase(ref.titulo));

        // Número do episódio (se houver)
        if (ref.numeroEpisodio) {
            partes.push(`(Nº ${ref.numeroEpisodio})`);
        }

        // Descritor
        partes.push('[Episódio de podcast].');

        // In + Nome do podcast
        partes.push(`In *${ref.nomePodcast}*.`);

        // Produtora (se houver)
        if (ref.produtora) {
            partes.push(`${ref.produtora}.`);
        }

        // URL
        if (ref.url) {
            partes.push(ref.url);
        }

        return partes.join(' ');
    }

    // ==========================================================================
    // FORMATAÇÃO DE AUTORES
    // ==========================================================================

    private formatarAutores(autores: Autor[], autorCorporativo?: { nomeCompleto: string; sigla?: string }): string {
        // Se for autor corporativo
        if (autorCorporativo) {
            return `${autorCorporativo.nomeCompleto}.`;
        }

        // Se não há autores
        if (!autores || autores.length === 0) {
            return '';
        }

        // Formatar cada autor
        const autoresFormatados: string[] = [];

        // Até 20 autores: listar todos
        if (autores.length <= 20) {
            autores.forEach((autor) => {
                const inicial = autor.inicialMeio
                    ? `${autor.nome.charAt(0)}. ${autor.inicialMeio}.`
                    : `${autor.nome.charAt(0)}.`;

                autoresFormatados.push(`${autor.sobrenome}, ${inicial}`);
            });
        } else {
            // Mais de 20 autores: 19 primeiros + ... + último
            for (let i = 0; i < 19; i++) {
                const autor = autores[i];
                if (autor) {
                    const inicial = autor.inicialMeio
                        ? `${autor.nome.charAt(0)}. ${autor.inicialMeio}.`
                        : `${autor.nome.charAt(0)}.`;

                    autoresFormatados.push(`${autor.sobrenome}, ${inicial}`);
                }
            }

            autoresFormatados.push('...');

            const ultimoAutor = autores[autores.length - 1];
            if (ultimoAutor) {
                const inicialUltimo = ultimoAutor.inicialMeio
                    ? `${ultimoAutor.nome.charAt(0)}. ${ultimoAutor.inicialMeio}.`
                    : `${ultimoAutor.nome.charAt(0)}.`;

                autoresFormatados.push(`${ultimoAutor.sobrenome}, ${inicialUltimo}`);
            }
        }

        // Juntar autores com vírgulas e '&' antes do último
        if (autoresFormatados.length === 1) {
            return `${autoresFormatados[0]}.`;
        } else if (autoresFormatados.length === 2) {
            return `${autoresFormatados[0]}, & ${autoresFormatados[1]}.`;
        } else {
            const todosExcetoUltimo = autoresFormatados.slice(0, -1).join(', ');
            const ultimo = autoresFormatados[autoresFormatados.length - 1];
            return `${todosExcetoUltimo}, & ${ultimo}.`;
        }
    }

    // ==========================================================================
    // LÓGICA DE ET AL. NAS CITAÇÕES
    // ==========================================================================

    private deveUsarEtAl(quantidadeAutores: number, primeiraCitacao: boolean = false): boolean {
        const limiar = this.config.usarEtAl.apartirDeQuantosAutores;

        // Se o limiar configurado for menor que 3 (padrão APA 7), respeita a config
        // APA 7 diz: 1 ou 2 autores = listar todos; 3 ou mais = et al.
        if (quantidadeAutores >= limiar) {
            if (primeiraCitacao) {
                return this.config.usarEtAl.primeiraVez;
            }
            return true;
        }
        return false;
    }

    private formatarEditores(editores: Autor[]): string {
        if (!editores || editores.length === 0) {
            return '';
        }

        const editoresFormatados: string[] = [];

        editores.forEach(editor => {
            const inicial = editor.inicialMeio
                ? `${editor.nome.charAt(0)}. ${editor.inicialMeio}.`
                : `${editor.nome.charAt(0)}.`;

            editoresFormatados.push(`${inicial} ${editor.sobrenome}`);
        });

        const sufixo = editores.length === 1 ? ABREVIACOES.ed : ABREVIACOES.eds;

        if (editoresFormatados.length === 1) {
            return `${editoresFormatados[0]} (${sufixo}),`;
        } else if (editoresFormatados.length === 2) {
            return `${editoresFormatados[0]} & ${editoresFormatados[1]} (${sufixo}),`;
        } else {
            const todosExcetoUltimo = editoresFormatados.slice(0, -1).join(', ');
            const ultimo = editoresFormatados[editoresFormatados.length - 1];
            return `${todosExcetoUltimo}, & ${ultimo} (${sufixo}),`;
        }
    }

    // ==========================================================================
    // GERAÇÃO DE CITAÇÕES
    // ==========================================================================

    private gerarCitacaoNarrativa(ref: Referencia): string {
        if (ref.autorCorporativo) {
            return ref.autorCorporativo.nomeCompleto;
        }

        if (!ref.autores || ref.autores.length === 0) {
            // Usar título entre aspas
            const tituloAbreviado = this.abreviarTitulo(ref.titulo);
            return `"${tituloAbreviado}"`;
        }

        const autores = ref.autores;
        const qtdAutores = autores.length;

        // Verificar se deve usar et al.
        // Assumindo que esta é uma geração genérica, tratamos como subsequente ou baseado na config
        const usarEtAl = this.deveUsarEtAl(qtdAutores, false);

        // Se deve usar et al., retorna imediatamente
        if (usarEtAl && autores[0]) {
            return `${autores[0].sobrenome} ${ABREVIACOES.et_al} (${ref.ano})`;
        }

        if (qtdAutores === 1 && autores[0]) {
            return `${autores[0].sobrenome} (${ref.ano})`;
        } else if (qtdAutores === 2 && autores[0] && autores[1]) {
            return `${autores[0].sobrenome} e ${autores[1].sobrenome} (${ref.ano})`;
        } else {
            // Caso fallback: mais de 2 autores mas configuração diz para não usar et al. (ex: threshold 5)
            // Neste caso, lista todos (ou até o limite se houver outra lógica)
            const nomes = autores.map((a: Autor) => a.sobrenome);
            const ultimo = nomes.pop();
            return `${nomes.join(', ')} e ${ultimo} (${ref.ano})`;
        }
    }

    private gerarCitacaoParentetica(ref: Referencia): string {
        if (ref.autorCorporativo) {
            return `(${ref.autorCorporativo.nomeCompleto}, ${ref.ano})`;
        }

        if (!ref.autores || ref.autores.length === 0) {
            const tituloAbreviado = this.abreviarTitulo(ref.titulo);
            return `("${tituloAbreviado}", ${ref.ano})`;
        }

        const autores = ref.autores;
        const qtdAutores = autores.length;
        const usarEtAl = this.deveUsarEtAl(qtdAutores, false);

        if (usarEtAl && autores[0]) {
            return `(${autores[0].sobrenome} ${ABREVIACOES.et_al}, ${ref.ano})`;
        }

        if (qtdAutores === 1 && autores[0]) {
            return `(${autores[0].sobrenome}, ${ref.ano})`;
        } else if (qtdAutores === 2 && autores[0] && autores[1]) {
            return `(${autores[0].sobrenome} & ${autores[1].sobrenome}, ${ref.ano})`;
        } else {
            // Fallback
            const nomes = autores.map((a: Autor) => a.sobrenome);
            const ultimo = nomes.pop();
            return `(${nomes.join(', ')} & ${ultimo}, ${ref.ano})`;
        }
    }

    // ==========================================================================
    // UTILIDADES DE FORMATAÇÃO
    // ==========================================================================

    private aplicarSentenceCase(texto: string): string {
        if (!texto) return '';

        // Dividir por dois pontos para tratar subtítulos
        const partes = texto.split(':');

        const partesFormatadas = partes.map((parte) => {
            const parteTrimmed = parte.trim();
            if (parteTrimmed.length === 0) return '';

            // Capitalizar primeira letra de cada parte (título e subtítulo)
            return parteTrimmed.charAt(0).toUpperCase() + parteTrimmed.slice(1).toLowerCase();
        });

        return partesFormatadas.join(': ');
    }

    private aplicarTitleCase(texto: string): string {
        if (!texto) return '';

        // Palavras menores que geralmente não são capitalizadas (exceto no início ou após pontuação)
        const palavrasMenores = ['a', 'an', 'the', 'and', 'but', 'or', 'for', 'nor', 'on', 'at', 'to', 'by', 'of', 'in', 'with', 'as'];

        // Tratamento básico: split por espaço
        // Para ser mais robusto, precisaria de regex para pontuação, mas vamos manter simples por enquanto
        const palavras = texto.split(' ');

        const palavrasFormatadas = palavras.map((palavra, index) => {
            // Se for a primeira palavra, sempre capitaliza
            if (index === 0) {
                return this.capitalizarPrimeira(palavra);
            }

            // Verificar se a palavra anterior termina com pontuação que exige capitalização (:, ., ?)
            const palavraAnterior = palavras[index - 1];
            if (palavraAnterior && /[:.?!]$/.test(palavraAnterior)) {
                return this.capitalizarPrimeira(palavra);
            }

            // Se for palavra menor, minúscula
            if (palavrasMenores.includes(palavra.toLowerCase())) {
                return palavra.toLowerCase();
            }

            // Caso padrão: capitaliza
            return this.capitalizarPrimeira(palavra);
        });

        return palavrasFormatadas.join(' ');
    }

    private capitalizarPrimeira(texto: string): string {
        return texto.charAt(0).toUpperCase() + texto.slice(1).toLowerCase();
    }

    private abreviarTitulo(titulo: string, maxPalavras: number = 3): string {
        const palavras = titulo.split(' ');
        return palavras.slice(0, maxPalavras).join(' ') + (palavras.length > maxPalavras ? '...' : '');
    }

    // ==========================================================================
    // VALIDAÇÃO
    // ==========================================================================

    private validar(ref: Referencia): ValidacaoReferencia {
        const camposObrigatoriosFaltantes: string[] = [];
        const camposOpcionaisFaltantes: string[] = [];
        const avisos: string[] = [];
        const erros: string[] = [];

        // Validações gerais
        if (!ref.titulo || ref.titulo.trim() === '') {
            camposObrigatoriosFaltantes.push('titulo');
            erros.push('Título é obrigatório');
        }

        if (!ref.ano || ref.ano.trim() === '') {
            camposObrigatoriosFaltantes.push('ano');
            erros.push('Ano é obrigatório');
        }

        if ((!ref.autores || ref.autores.length === 0) && !ref.autorCorporativo) {
            avisos.push('Nenhum autor especificado - usar título como identificador');
        }

        // Validações específicas por tipo
        if (isReferenciaArtigo(ref)) {
            if (!ref.periodico) camposObrigatoriosFaltantes.push('periodico');
            if (!ref.volume) camposObrigatoriosFaltantes.push('volume');
            if (!ref.paginas) camposObrigatoriosFaltantes.push('paginas');
            if (!ref.doi && !ref.url) {
                avisos.push('DOI ou URL recomendado para artigos');
            }

            // Validação básica de URL se presente
            if (ref.url && !ref.url.startsWith('http')) {
                avisos.push('URL deve começar com http:// ou https://');
            }
        }

        if (isReferenciaLivro(ref)) {
            if (!ref.editora) camposObrigatoriosFaltantes.push('editora');
        }

        const valida = erros.length === 0;
        const completa = camposObrigatoriosFaltantes.length === 0 && avisos.length === 0;

        return {
            valida,
            completa,
            camposObrigatoriosFaltantes,
            camposOpcionaisFaltantes,
            avisos,
            erros,
        };
    }

    // ==========================================================================
    // GERAÇÃO DE METADADOS
    // ==========================================================================

    private gerarMetadados(ref: Referencia): MetadadosReferencia {
        return {
            temDOI: !!ref.doi,
            temURL: !!ref.url,
            quantidadeAutores: ref.autores?.length || 0,
            contemAutorCorporativo: !!ref.autorCorporativo,
            idioma: ref.idioma || this.config.idiomaPadrao,
        };
    }
}
