/**
 * APA Quote Formatter - Formatador de Aspas para Normas APA 7ª Edição
 * 
 * Este script automatiza a formatação de aspas em documentos acadêmicos
 * seguindo as diretrizes da American Psychological Association (7ª edição)
 * 
 * @author Programador Sênior TypeScript
 * @version 1.0.0
 */

/// <reference types="office-js" />

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
        private undoStack: Array<{ range: Word.Range; text: string }> = [];
        private paragraphOffsets: number[] = [];

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
                    body.load('text');
                    await context.sync();

                    // Obter range de trabalho (seleção ou documento inteiro)
                    // Cast to any to allow both Range and Body in helper methods
                    const workingRange = (this.config.applyToSelection
                        ? context.document.getSelection()
                        : body) as any;

                    workingRange.load('text, paragraphs');
                    await context.sync();

                    // 1. Converter aspas retas em tipográficas
                    if (this.config.convertQuotes) {
                        await this.convertStraightToTypographicQuotes(context, workingRange);
                    }

                    // 2. Validar uso de aspas
                    if (this.config.validateUsage) {
                        await this.validateQuoteUsage(context, workingRange);
                    }

                    // 3. Corrigir pontuação
                    if (this.config.fixPunctuation) {
                        await this.fixPunctuation(context, workingRange);
                    }

                    // 4. Identificar citações longas
                    if (this.config.identifyLongQuotes) {
                        await this.identifyLongQuotes(context, workingRange);
                    }

                    await context.sync();
                });

                return this.report;
            } catch (error) {
                console.error('Erro durante execução:', error);
                throw new Error(`Erro ao formatar documento: ${(error as Error).message}`);
            }
        }

        /**
         * Converte aspas retas em aspas tipográficas curvas
         */
        private async convertStraightToTypographicQuotes(
            context: Word.RequestContext,
            range: Word.Range
        ): Promise<void> {
            range.load('text');
            await context.sync();

            const text = range.text;
            let newText = text;
            let conversionCount = 0;

            // Padrão para identificar aspas duplas retas
            // Converte " no início de palavra ou após espaço em "
            newText = newText.replace(QuoteFormatter.Patterns.DoubleQuoteStraightOpen, (_match, before, after) => {
                conversionCount++;
                return before + '"' + after;
            });

            // Converte " no final de palavra ou antes de espaço/pontuação em "
            newText = newText.replace(QuoteFormatter.Patterns.DoubleQuoteStraightClose, (_match, before) => {
                conversionCount++;
                return before + '"';
            });

            // Converte aspas simples para contextos específicos
            // Mantém aspas simples apenas em citações dentro de citações
            newText = this.handleSingleQuotes(newText);

            if (newText !== text) {
                // Salvar para desfazer
                this.undoStack.push({ range: range, text: text });

                // Aplicar mudanças
                range.insertText(newText, Word.InsertLocation.replace);
                this.report.quotesConverted += conversionCount;
                this.report.straightToTypographic += conversionCount;

                await context.sync();
            }
        }

        /**
         * Trata aspas simples, convertendo apenas quando apropriado
         */
        private handleSingleQuotes(text: string): string {
            // Identifica citações dentro de citações: "texto 'citação interna' texto"

            // Primeiro, protege aspas simples em citações aninhadas
            let result = text.replace(QuoteFormatter.Patterns.NestedQuote, (match) => {
                // Mantém aspas simples dentro de duplas
                return match.replace(/'/g, '’'); // Aspas simples tipográficas
            });

            // Converte aspas simples isoladas (incorretas) em duplas
            // Apenas se não estiverem dentro de aspas duplas
            result = result.replace(QuoteFormatter.Patterns.IncorrectSingleQuote, (_match, before, content) => {
                this.report.singleToDouble++;
                return before + '"' + content + '"';
            });

            return result;
        }

        /**
         * Valida o uso de aspas segundo normas APA
         */
        private async validateQuoteUsage(
            context: Word.RequestContext,
            range: Word.Range
        ): Promise<void> {
            // [OPTIMIZED] Load all paragraphs data at once to avoid sync in loop
            const paragraphs = range.paragraphs;
            paragraphs.load('items/text'); // Load text property for all items
            await context.sync();

            // Perform validation in memory
            for (let i = 0; i < paragraphs.items.length; i++) {
                const paragraph = paragraphs.items[i];
                if (!paragraph) continue;
                const text = paragraph.text;

                if (!text) continue;

                // Verificar termos técnicos entre aspas que deveriam estar em itálico
                this.checkTechnicalTerms(text, i + 1);

                // Verificar âncoras de escala
                this.checkScaleAnchors(text, i + 1);

                // Verificar uso incorreto de aspas simples
                this.checkIncorrectSingleQuotes(text, i + 1);
            }
        }

        /**
         * Verificar termos técnicos que deveriam estar em itálico
         */
        private checkTechnicalTerms(text: string, paragraphNumber: number): void {
            QuoteFormatter.Patterns.TechnicalTerms.forEach(pattern => {
                const matches = text.match(pattern);
                if (matches) {
                    matches.forEach(match => {
                        this.report.issues.push({
                            type: 'technical_term',
                            location: `Parágrafo ${paragraphNumber}`,
                            text: match,
                            suggestion: `Considere usar itálico em vez de aspas: ${match.replace(/"/g, '')}`,
                            paragraph: paragraphNumber
                        });
                    });
                }
            });
        }

        /**
         * Verifica âncoras de escala entre aspas
         */
        private checkScaleAnchors(text: string, paragraphNumber: number): void {
            const matches = text.matchAll(QuoteFormatter.Patterns.ScaleAnchor);
            for (const match of matches) {
                this.report.issues.push({
                    type: 'scale_anchor',
                    location: `Parágrafo ${paragraphNumber}`,
                    text: match[0],
                    suggestion: `Âncoras de escala devem usar itálico: ${match[1]} = ${match[2]}`,
                    paragraph: paragraphNumber
                });
            }
        }

        /**
         * Verifica uso incorreto de aspas simples
         */
        private checkIncorrectSingleQuotes(text: string, paragraphNumber: number): void {
            const matches = text.matchAll(QuoteFormatter.Patterns.IncorrectSingleQuoteContext);
            for (const match of matches) {
                // Verificar se não está dentro de aspas duplas
                const beforeQuote = text.substring(0, match.index!);
                // Simple check for balanced double quotes before this point
                const openDoubleQuotes = (beforeQuote.match(/"/g) || []).length;

                if (openDoubleQuotes % 2 === 0) {
                    this.report.issues.push({
                        type: 'wrong_single_quote',
                        location: `Parágrafo ${paragraphNumber}`,
                        text: match[0],
                        suggestion: `Use aspas duplas em vez de simples: "${match[1]}"`,
                        paragraph: paragraphNumber
                    });
                }
            }
        }

        /**
         * Corrige pontuação em relação às aspas
         */
        private async fixPunctuation(
            context: Word.RequestContext,
            range: Word.Range
        ): Promise<void> {
            range.load('text');
            await context.sync();

            let text = range.text;
            const originalText = text;
            let fixCount = 0;

            // Regra APA: pontos e vírgulas dentro das aspas
            // "texto".<space> -> "texto."<space>
            text = text.replace(QuoteFormatter.Patterns.PunctuationPeriod, (_match) => {
                fixCount++;
                return '."';
            });

            text = text.replace(QuoteFormatter.Patterns.PunctuationComma, (_match) => {
                fixCount++;
                return ',"';
            });

            // Regra APA: ponto-e-vírgula e dois-pontos fora das aspas
            text = text.replace(QuoteFormatter.Patterns.PunctuationSemiColon, '"; ');
            text = text.replace(QuoteFormatter.Patterns.PunctuationColon, '": ');

            // Corrigir pontuação com referência parentética
            // "texto" (Autor, 2020). -> "texto" (Autor, 2020).
            text = text.replace(QuoteFormatter.Patterns.ParentheticalRef, '" ($2).');

            if (text !== originalText) {
                this.undoStack.push({ range: range, text: originalText });
                range.insertText(text, Word.InsertLocation.replace);
                this.report.punctuationFixed = fixCount;
                await context.sync();
            }
        }

        /**
         * Identifica citações longas (40+ palavras) que devem ser formatadas como bloco
         */
        private async identifyLongQuotes(
            context: Word.RequestContext,
            range: Word.Range
        ): Promise<void> {
            range.load('text');
            await context.sync();

            const text = range.text;

            // Encontrar todas as citações entre aspas
            const matches = text.matchAll(QuoteFormatter.Patterns.LongQuote); // Mínimo de 100 caracteres para verificar

            // Build paragraph offsets cache once
            this.buildParagraphOffsets(text);

            for (const match of matches) {
                const quotedText = match[1];
                const wordCount = this.countWords(quotedText || '');

                if (wordCount >= 40) {
                    // Encontrar parágrafo
                    const position = match.index!;
                    const paragraphNumber = this.findParagraphNumber(text, position);

                    this.report.longQuotesFound++;
                    this.report.issues.push({
                        type: 'long_quote',
                        location: `Parágrafo ${paragraphNumber}`,
                        text: match[0].substring(0, 100) + '...',
                        suggestion: `Esta citação tem ${wordCount} palavras e deve ser formatada como bloco (sem aspas, recuo de 1,27 cm)`,
                        paragraph: paragraphNumber
                    });
                }
            }
        }

        /**
         * Converte citação longa em formato de bloco
         */
        public async convertToBlockQuote(
            context: Word.RequestContext,
            paragraphNumber: number
        ): Promise<void> {
            const paragraphs = context.document.body.paragraphs;
            paragraphs.load('items');
            await context.sync();

            if (paragraphNumber > 0 && paragraphNumber <= paragraphs.items.length) {
                const paragraph = paragraphs.items[paragraphNumber - 1];
                if (!paragraph) return;

                paragraph.load('text, leftIndent, firstLineIndent');
                await context.sync();

                let text = paragraph.text;
                if (!text) return;

                // Remover aspas
                text = text.replace(/^"/, '').replace(/"$/, '');

                // Aplicar formatação
                paragraph.leftIndent = 36; // 1,27 cm em pontos (aproximadamente)
                paragraph.firstLineIndent = 0;
                paragraph.insertText(text, Word.InsertLocation.replace);

                await context.sync();
            }
        }

        /**
         * Conta palavras em um texto
         */
        private countWords(text: string): number {
            return text.trim().split(/\s+/).length;
        }

        /**
         * Constrói cache de offsets de parágrafos
         */
        private buildParagraphOffsets(text: string): void {
            this.paragraphOffsets = [];
            let currentPos = 0;
            // Split once to find lengths
            // Note: This assumes \n is the paragraph delimiter in the text property
            const paragraphs = text.split('\n');
            for (const p of paragraphs) {
                currentPos += p.length + 1; // +1 for the newline char (or delimiter)
                this.paragraphOffsets.push(currentPos); // End position of this paragraph
            }
        }

        /**
         * Encontra o número do parágrafo baseado na posição no texto (Optimized with binary search or simple cache lookup)
         */
        private findParagraphNumber(text: string, position: number): number {
            if (this.paragraphOffsets.length === 0) {
                // Fallback if cache not built
                const textUpToPosition = text.substring(0, position);
                return textUpToPosition.split('\n').length;
            }

            // Find the first offset strictly greater than position
            // Since paragraphs order is sequential, we can just findIndex or binary search
            // For simplicity in JS:
            const index = this.paragraphOffsets.findIndex(offset => offset > position);
            return index !== -1 ? index + 1 : this.paragraphOffsets.length;
        }

        /**
         * Desfaz todas as alterações realizadas
         */
        public async undoChanges(context: Word.RequestContext): Promise<void> {
            for (const item of this.undoStack.reverse()) {
                item.range.insertText(item.text, Word.InsertLocation.replace);
            }
            await context.sync();
            this.undoStack = [];
        }

        /**
         * Gera relatório textual das operações realizadas
         */
        public generateTextReport(): string {
            let report = '=== RELATÓRIO DE FORMATAÇÃO APA ===\n\n';

            report += `📊 ESTATÍSTICAS:\n`;
            report += `• Total de aspas convertidas: ${this.report.quotesConverted}\n`;
            report += `• Aspas retas → tipográficas: ${this.report.straightToTypographic}\n`;
            report += `• Aspas simples → duplas: ${this.report.singleToDouble}\n`;
            report += `• Pontuações corrigidas: ${this.report.punctuationFixed}\n`;
            report += `• Citações longas encontradas: ${this.report.longQuotesFound}\n\n`;

            if (this.report.issues.length > 0) {
                report += `⚠️ PROBLEMAS ENCONTRADOS (${this.report.issues.length}):\n\n`;

                const issuesByType = this.groupIssuesByType();

                for (const [type, issues] of Object.entries(issuesByType)) {
                    report += `\n${this.getIssueTypeLabel(type)}:\n`;
                    issues.forEach((issue, index) => {
                        report += `  ${index + 1}. ${issue.location}\n`;
                        report += `     Texto: ${issue.text}\n`;
                        report += `     💡 ${issue.suggestion}\n\n`;
                    });
                }
            } else {
                report += '✅ Nenhum problema encontrado!\n';
            }

            report += '\n=== FIM DO RELATÓRIO ===';
            return report;
        }

        /**
         * Agrupa problemas por tipo
         */
        private groupIssuesByType(): Record<string, ValidationIssue[]> {
            const grouped: Record<string, ValidationIssue[]> = {};

            this.report.issues.forEach(issue => {
                if (!grouped[issue.type]) {
                    grouped[issue.type] = [];
                }
                grouped[issue.type]!.push(issue);
            });

            return grouped;
        }

        /**
         * Retorna label descritivo para tipo de problema
         */
        private getIssueTypeLabel(type: string): string {
            const labels: Record<string, string> = {
                'long_quote': '📏 Citações Longas (devem ser formatadas como bloco)',
                'technical_term': '📚 Termos Técnicos (considere usar itálico)',
                'wrong_single_quote': '❌ Aspas Simples Incorretas',
                'punctuation': '🔤 Problemas de Pontuação',
                'scale_anchor': '📊 Âncoras de Escala (devem usar itálico)'
            };
            return labels[type] || type;
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
   • Mantém aspas simples apenas em citações dentro de citações

2️⃣ VALIDAR USO
   • Identifica termos técnicos que deveriam estar em itálico
   • Detecta uso incorreto de aspas simples
   • Verifica âncoras de escala

3️⃣ CORRIGIR PONTUAÇÃO
   • Coloca pontos e vírgulas dentro das aspas
   • Mantém ponto-e-vírgula e dois-pontos fora das aspas
   • Corrige pontuação com referências parentéticas

4️⃣ IDENTIFICAR CITAÇÕES LONGAS
   • Detecta citações com 40+ palavras
   • Sugere conversão para formato de bloco

REGRAS APA APLICADAS:

✓ Títulos de artigos/capítulos: entre aspas duplas
✓ Citações diretas < 40 palavras: entre aspas duplas
✓ Citações ≥ 40 palavras: bloco recuado sem aspas
✓ Citação dentro de citação: aspas simples internas
✓ Termos técnicos: itálico (não aspas)
✓ Ironia/ênfase: aspas duplas (primeira ocorrência)

COMO USAR:

1. Selecione as opções desejadas
2. Escolha aplicar ao documento inteiro ou apenas à seleção
3. Clique em "Executar"
4. Revise o relatório gerado
5. Use "Desfazer" (Ctrl+Z) se necessário

Para mais informações sobre normas APA, consulte:
https://apastyle.apa.org/
        `;

        alert(helpText);
    }
}

// Expor funções globalmente para uso no HTML
(window as any).APAQuoteFormatter = APAQuoteFormatter;

// Inicializar quando o Office estiver pronto
Office.onReady((info) => {
    if (info.host === Office.HostType.Word) {
        console.log('APA Quote Formatter carregado com sucesso!');
    }
});
