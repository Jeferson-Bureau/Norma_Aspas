/**
 * APA Quote Formatter - Formatador de Aspas para Normas APA 7ª Edição
 * 
 * Este script automatiza a formatação de aspas em documentos acadêmicos
 * seguindo as diretrizes da American Psychological Association (7ª edição)
 * 
 * @author Programador Sênior TypeScript
 * @version 1.1.0
 */

/// <reference types="office-js" />

import { FormatadorAPA7 } from './references/FormatadorAPA7';

namespace APAQuoteFormatter {

    /**
     * Interface para configurações do formatador
     */
    interface FormatterConfig {
        convertQuotes: boolean;
        validateUsage: boolean;
        fixPunctuation: boolean;
        identifyLongQuotes: boolean;
        generateReport: boolean;
        applyToSelection: boolean;
    }

    /**
     * Namespace para geração de referências
     */
    export namespace ReferenceGenerator {
        let lastGeneratedReference: string = '';

        export function generateReference(data: any): any {
            const formatador = new FormatadorAPA7({
                idiomaPadrao: 'pt',
                incluirDOI: true,
                validacaoEstrita: false // Permitir geração mesmo com dados parciais para preview
            });

            try {
                // Mapear dados do formulário para o tipo Referencia
                // Simplificação: assumindo que o formulário envia dados compatíveis ou faremos o mapeamento aqui
                // Para este exemplo, vamos supor que 'data' já vem estruturado ou faremos um mapeamento básico

                // Exemplo de mapeamento básico para Artigo
                let referencia: any = { ...data };

                // Limpar campos vazios
                Object.keys(referencia).forEach(key => {
                    if (referencia[key] === '' || referencia[key] === null || referencia[key] === undefined) {
                        delete referencia[key];
                    }
                });

                // Tratamento de autores se vierem como string (ex: "Silva, J.; Santos, M.")
                if (typeof referencia.autores === 'string') {
                    // Lógica de parsing simplificada ou esperar JSON
                    // Aqui vamos assumir que a UI envia o objeto correto ou vamos implementar um parser simples
                    // Por enquanto, vamos assumir que a UI envia autores estruturados ou vamos deixar vazio
                }

                const resultado = formatador.formatar(referencia);
                lastGeneratedReference = resultado.referenciaCompleta;

                return {
                    success: true,
                    reference: resultado.referenciaCompleta,
                    citation: resultado.citacaoParentetica,
                    narrative: resultado.citacaoNarrativa
                };
            } catch (error) {
                return {
                    success: false,
                    error: (error as Error).message
                };
            }
        }

        export async function insertReference(): Promise<void> {
            if (!lastGeneratedReference) return;

            await Word.run(async (context) => {
                const selection = context.document.getSelection();
                selection.insertText(lastGeneratedReference, Word.InsertLocation.replace);
                // Adicionar quebra de linha após
                selection.insertParagraph("", Word.InsertLocation.after);
                await context.sync();
            });
        }
    }

    /**
     * Interface para problemas encontrados
     */
    interface ValidationIssue {
        type: 'long_quote' | 'technical_term' | 'wrong_single_quote' | 'punctuation' | 'scale_anchor';
        location: string;
        text: string;
        suggestion: string;
        paragraph: number;
    }

    /**
     * Interface para relatório de execução
     */
    interface ExecutionReport {
        quotesConverted: number;
        straightToTypographic: number;
        singleToDouble: number;
        issues: ValidationIssue[];
        punctuationFixed: number;
        longQuotesFound: number;
    }

    /**
     * Classe principal do formatador APA
     */
    export class QuoteFormatter {
        private static readonly Patterns = {
            DoubleQuoteStraightOpen: /(^|[\s\(])"([^"])/g,
            DoubleQuoteStraightClose: /([^"])"/g,
            NestedQuote: /"([^"]*)'([^']+)'([^"]*)"/g,
            IncorrectSingleQuote: /(?<!")(\s|^)'([^']+)'(?!")/g,
            IncorrectSingleQuoteContext: /(?<!")\s'([^']+)'(?!")/g,
            TechnicalTerms: [
                /"[a-z]+-[a-z]+"/gi,
                /"(?:teste|escala|modelo|teoria|método)\s+(?:de|do|da)\s+\w+"/gi,
            ],
            ScaleAnchor: /"(\d+)\s*=\s*([^"]+)"/g,
            LongQuote: /"([^"]{100,})"/g,
            PunctuationPeriod: /(")\s*\./g,
            PunctuationComma: /(")\s*,/g,
            PunctuationSemiColon: /;("\s*)/g,
            PunctuationColon: /:("\s*)/g,
            ParentheticalRef: /"(\s*)\(([^)]+)\)\s*\./g
        };

        private config: FormatterConfig;
        private report: ExecutionReport;

        constructor(config: FormatterConfig) {
            this.config = config;
            this.report = {
                quotesConverted: 0,
                straightToTypographic: 0,
                singleToDouble: 0,
                issues: [],
                punctuationFixed: 0,
                longQuotesFound: 0
            };
        }

        /**
         * Executa o processo de formatação completo
         */
        public async execute(): Promise<ExecutionReport> {
            try {
                await Word.run(async (context) => {
                    const body = context.document.body;

                    // Obter range de trabalho (seleção ou documento inteiro)
                    // Cast to any to allow both Range and Body in helper methods
                    const workingRange = (this.config.applyToSelection
                        ? context.document.getSelection()
                        : body) as any;

                    // Carregar parágrafos para iteração
                    const paragraphs = workingRange.paragraphs;
                    paragraphs.load('items');
                    await context.sync();

                    // Iterar por parágrafos para evitar substituição destrutiva de todo o corpo
                    for (let i = 0; i < paragraphs.items.length; i++) {
                        const paragraph = paragraphs.items[i];
                        if (!paragraph) continue;

                        // 1. Converter aspas retas em tipográficas
                        if (this.config.convertQuotes) {
                            await this.convertQuotesInParagraph(context, paragraph);
                        }

                        // 2. Corrigir pontuação
                        if (this.config.fixPunctuation) {
                            await this.fixPunctuationInParagraph(context, paragraph);
                        }
                    }

                    // Validações (apenas leitura)
                    if (this.config.validateUsage || this.config.identifyLongQuotes) {
                        // Cast workingRange again or ensure performValidations accepts 'any' or intersection
                        await this.performValidations(context, workingRange);
                    }
                });

                return this.report;
            } catch (error) {
                console.error('Erro durante execução:', error);
                throw new Error(`Erro ao formatar documento: ${(error as Error).message}`);
            }
        }

        /**
         * Converte aspas em um parágrafo específico preservando formatação
         */
        private async convertQuotesInParagraph(context: Word.RequestContext, paragraph: Word.Paragraph): Promise<void> {
            paragraph.load('text');
            await context.sync();
            const text = paragraph.text;
            if (!text) return;

            // Converter aspas duplas de abertura: "Texto -> “Texto
            const openMatches = Array.from(text.matchAll(QuoteFormatter.Patterns.DoubleQuoteStraightOpen));
            for (const match of openMatches) {
                const searchResults = paragraph.search(match[0], { matchCase: true });
                searchResults.load('items');
                await context.sync();

                if (!searchResults.items) continue;

                for (const range of searchResults.items) {
                    const quoteRanges = range.search('"', { matchCase: true });
                    quoteRanges.load('items');
                    await context.sync();

                    if (quoteRanges.items && quoteRanges.items.length > 0) {
                        const quote = quoteRanges.items[0];
                        if (quote) {
                            quote.insertText('“', Word.InsertLocation.replace);
                            this.report.quotesConverted++;
                            this.report.straightToTypographic++;
                        }
                    }
                }
            }

            if (openMatches.length > 0) await context.sync();

            // Converter aspas duplas de fechamento: Texto" -> Texto”
            paragraph.load('text');
            await context.sync();
            const updatedText = paragraph.text;

            const closeMatches = Array.from(updatedText.matchAll(QuoteFormatter.Patterns.DoubleQuoteStraightClose));
            for (const match of closeMatches) {
                const searchResults = paragraph.search(match[0], { matchCase: true });
                searchResults.load('items');
                await context.sync();

                if (!searchResults.items) continue;

                for (const range of searchResults.items) {
                    const quoteRanges = range.search('"', { matchCase: true });
                    quoteRanges.load('items');
                    await context.sync();

                    if (quoteRanges.items && quoteRanges.items.length > 0) {
                        const quote = quoteRanges.items[0];
                        if (quote) {
                            quote.insertText('”', Word.InsertLocation.replace);
                            this.report.quotesConverted++;
                            this.report.straightToTypographic++;
                        }
                    }
                }
            }
        }

        /**
         * Corrige pontuação em um parágrafo preservando formatação
         */
        private async fixPunctuationInParagraph(context: Word.RequestContext, paragraph: Word.Paragraph): Promise<void> {
            paragraph.load('text');
            await context.sync();
            const text = paragraph.text;
            if (!text) return;

            // Ponto: ". -> ."
            const periodMatches = Array.from(text.matchAll(QuoteFormatter.Patterns.PunctuationPeriod));
            for (const match of periodMatches) {
                const searchResults = paragraph.search(match[0], { matchCase: true });
                searchResults.load('items');
                await context.sync();

                for (const range of searchResults.items) {
                    range.insertText('.”', Word.InsertLocation.replace);
                    this.report.punctuationFixed++;
                }
            }

            // Vírgula: ", -> ,"
            const commaMatches = Array.from(text.matchAll(QuoteFormatter.Patterns.PunctuationComma));
            for (const match of commaMatches) {
                const searchResults = paragraph.search(match[0], { matchCase: true });
                searchResults.load('items');
                await context.sync();

                for (const range of searchResults.items) {
                    range.insertText(',”', Word.InsertLocation.replace);
                    this.report.punctuationFixed++;
                }
            }
        }

        /**
         * Realiza validações de uso (apenas leitura)
         */
        private async performValidations(context: Word.RequestContext, range: Word.Range): Promise<void> {
            range.load("text, paragraphs");
            await context.sync();

            const paragraphs = range.paragraphs;
            paragraphs.load('items/text');
            await context.sync();

            for (let i = 0; i < paragraphs.items.length; i++) {
                const p = paragraphs.items[i];
                if (!p || !p.text) continue;

                if (this.config.validateUsage) {
                    this.checkTechnicalTerms(p.text, i + 1);
                    this.checkScaleAnchors(p.text, i + 1);
                    this.checkIncorrectSingleQuotes(p.text, i + 1);
                }

                if (this.config.identifyLongQuotes) {
                    this.checkLongQuotes(p.text, i + 1);
                }
            }
        }

        private checkTechnicalTerms(text: string | undefined, pNum: number): void {
            if (!text) return;
            QuoteFormatter.Patterns.TechnicalTerms.forEach(pattern => {
                const matches = text.match(pattern);
                if (matches) {
                    matches.forEach(match => {
                        this.report.issues.push({
                            type: 'technical_term',
                            location: `Parágrafo ${pNum}`,
                            text: match,
                            suggestion: `Considere usar itálico: ${match.replace(/"/g, '')}`,
                            paragraph: pNum
                        });
                    });
                }
            });
        }

        private checkScaleAnchors(text: string | undefined, pNum: number): void {
            if (!text) return;
            const matches = text.matchAll(QuoteFormatter.Patterns.ScaleAnchor);
            for (const match of matches) {
                this.report.issues.push({
                    type: 'scale_anchor',
                    location: `Parágrafo ${pNum}`,
                    text: match[0],
                    suggestion: `Âncoras de escala devem usar itálico: ${match[1]} = ${match[2]}`,
                    paragraph: pNum
                });
            }
        }

        private checkIncorrectSingleQuotes(text: string | undefined, pNum: number): void {
            if (!text) return;
            const matches = text.matchAll(QuoteFormatter.Patterns.IncorrectSingleQuoteContext);
            for (const match of matches) {
                const beforeQuote = text.substring(0, match.index!);
                const openDoubleQuotes = (beforeQuote.match(/"/g) || []).length;

                if (openDoubleQuotes % 2 === 0) {
                    this.report.issues.push({
                        type: 'wrong_single_quote',
                        location: `Parágrafo ${pNum}`,
                        text: match[0],
                        suggestion: `Use aspas duplas em vez de simples: "${match[1]}"`,
                        paragraph: pNum
                    });
                }
            }
        }

        private checkLongQuotes(text: string | undefined, pNum: number): void {
            if (!text) return;
            const matches = text.matchAll(QuoteFormatter.Patterns.LongQuote);
            for (const match of matches) {
                if (match[1] && this.countWords(match[1]) >= 40) {
                    this.report.longQuotesFound++;
                    this.report.issues.push({
                        type: 'long_quote',
                        location: `Parágrafo ${pNum}`,
                        text: match[0].substring(0, 50) + '...',
                        suggestion: `Citação longa (40+ palavras) deve ser bloco`,
                        paragraph: pNum
                    });
                }
            }
        }

        private countWords(text: string): number {
            return text.trim().split(/\s+/).length;
        }

        public generateTextReport(): string {
            let report = '=== RELATÓRIO DE FORMATAÇÃO APA ===\n\n';
            report += `📊 ESTATÍSTICAS:\n`;
            report += `• Aspas convertidas: ${this.report.quotesConverted}\n`;
            report += `• Pontuações corrigidas: ${this.report.punctuationFixed}\n`;
            report += `• Problemas encontrados: ${this.report.issues.length}\n\n`;

            if (this.report.issues.length > 0) {
                report += `⚠️ DETALHES:\n`;
                this.report.issues.slice(0, 10).forEach(issue => {
                    report += `- [${issue.location}] ${issue.suggestion}\n`;
                });
                if (this.report.issues.length > 10) report += `...e mais ${this.report.issues.length - 10} itens.`;
            } else {
                report += '✅ Tudo certo!';
            }
            return report;
        }

    }

    /**
     * Inicializa a interface do usuário
     */
    export async function showUI(): Promise<void> {
        const dialog = document.getElementById('formatterDialog');
        if (dialog) {
            dialog.style.display = 'block';
        }
    }

    /**
     * Executa o formatador com as configurações selecionadas
     */
    export async function runFormatter(): Promise<void> {
        const config: FormatterConfig = {
            convertQuotes: (document.getElementById('convertQuotes') as HTMLInputElement)?.checked ?? true,
            validateUsage: (document.getElementById('validateUsage') as HTMLInputElement)?.checked ?? true,
            fixPunctuation: (document.getElementById('fixPunctuation') as HTMLInputElement)?.checked ?? true,
            identifyLongQuotes: (document.getElementById('identifyLongQuotes') as HTMLInputElement)?.checked ?? true,
            generateReport: (document.getElementById('generateReport') as HTMLInputElement)?.checked ?? true,
            applyToSelection: (document.getElementById('applyToSelection') as HTMLInputElement)?.checked ?? false
        };

        const statusDiv = document.getElementById('status');
        if (statusDiv) {
            statusDiv.textContent = 'Processando...';
            statusDiv.className = 'status processing';
        }

        try {
            const formatter = new QuoteFormatter(config);
            const report = await formatter.execute();

            if (config.generateReport) {
                const reportText = formatter.generateTextReport();
                showReport(reportText);
            }

            if (statusDiv) {
                statusDiv.textContent = `✅ Concluído! ${report.quotesConverted} aspas convertidas.`;
                statusDiv.className = 'status success';
            }
        } catch (error) {
            if (statusDiv) {
                statusDiv.textContent = `❌ Erro: ${(error as Error).message}`;
                statusDiv.className = 'status error';
            }
            console.error('Erro ao executar formatador:', error);
        }
    }

    /**
     * Exibe o relatório em uma janela modal
     */
    function showReport(reportText: string): void {
        const reportDiv = document.getElementById('reportContent');
        if (reportDiv) {
            reportDiv.textContent = reportText;
            const reportModal = document.getElementById('reportModal');
            if (reportModal) {
                reportModal.style.display = 'block';
            }
        }
    }

    /**
     * Fecha o modal de relatório
     */
    export function closeReport(): void {
        const reportModal = document.getElementById('reportModal');
        if (reportModal) {
            reportModal.style.display = 'none';
        }
    }

    /**
     * Exibe ajuda sobre o formatador
     */
    export function showHelp(): void {
        const helpText = `
FORMATADOR DE ASPAS APA 7ª EDIÇÃO

Este add-in automatiza a formatação de aspas seguindo as normas da American Psychological Association (7ª edição).

FUNCIONALIDADES:

1️⃣ CONVERTER ASPAS
   • Transforma aspas retas (" ") em aspas tipográficas curvas (" ")
   • Mantém a formatação original do texto (negrito, itálico, etc.)

2️⃣ VALIDAR USO
   • Identifica termos técnicos que deveriam estar em itálico
   • Detecta uso incorreto de aspas simples
   • Verifica âncoras de escala

3️⃣ CORRIGIR PONTUAÇÃO
   • Coloca pontos e vírgulas dentro das aspas

4️⃣ IDENTIFICAR CITAÇÕES LONGAS
   • Detecta citações com 40+ palavras

COMO USAR:
1. Selecione as opções desejadas
2. Escolha aplicar ao documento inteiro ou apenas à seleção
3. Clique em "Executar"
        `;

        alert(helpText);
    }
}

// Expor funções globalmente para uso no HTML
(window as any).APAQuoteFormatter = APAQuoteFormatter;
(window as any).ReferenceGenerator = APAQuoteFormatter.ReferenceGenerator;

// Inicializar quando o Office estiver pronto
Office.onReady((info) => {
    if (info.host === Office.HostType.Word) {
        console.log('APA Quote Formatter carregado com sucesso!');
    }
});
