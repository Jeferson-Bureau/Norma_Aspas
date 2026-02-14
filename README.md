# 📝 Formatador APA - Aspas para Microsoft Word

## Sobre o Projeto

Add-in profissional para Microsoft Word que automatiza a formatação de aspas em documentos acadêmicos seguindo rigorosamente as normas da **American Psychological Association (APA) 7ª edição**.

Desenvolvido com TypeScript, este add-in oferece funcionalidades avançadas para garantir que seu documento acadêmico esteja em conformidade com os padrões internacionais de formatação.

---

## ✨ Funcionalidades Principais

### 1️⃣ Conversão Automática de Aspas
- ✅ Converte aspas retas (`" "`) em aspas tipográficas curvas (`" "`)
- ✅ Identifica e mantém aspas simples (`' '`) apenas em contextos apropriados
- ✅ Gerencia automaticamente citações dentro de citações
- ✅ Converte aspas simples incorretas em aspas duplas quando necessário

### 2️⃣ Validação Segundo Normas APA
- 🔍 Detecta citações longas (40+ palavras) entre aspas que devem ser blocos
- 🔍 Identifica termos técnicos incorretamente entre aspas (devem estar em itálico)
- 🔍 Verifica uso incorreto de aspas simples fora de citações aninhadas
- 🔍 Detecta âncoras de escala entre aspas (devem usar itálico)

### 3️⃣ Formatação de Citações em Bloco
- 📐 Detecta automaticamente citações com 40+ palavras
- 📐 Oferece conversão automática para formato de bloco:
  - Remove aspas
  - Aplica recuo de 1,27 cm da margem esquerda
  - Mantém espaçamento adequado

### 4️⃣ Correção de Pontuação
- ✏️ Posiciona pontos finais e vírgulas **dentro** das aspas de fechamento
- ✏️ Mantém ponto-e-vírgula e dois-pontos **fora** das aspas
- ✏️ Corrige posicionamento quando há referência parentética após citação

### 5️⃣ Interface Intuitiva
- 🎨 Painel lateral moderno e responsivo
- ⚙️ Opções configuráveis para cada funcionalidade
- 📊 Relatório detalhado de alterações realizadas
- 🔄 Suporte para aplicar ao documento inteiro ou apenas à seleção

### 6️⃣ Relatório de Validação
- 📈 Número de aspas convertidas
- 📍 Lista de problemas com localização exata (número do parágrafo)
- 💡 Sugestões de correção manual
- 📋 Exportação de relatório detalhado

---

## 📋 Regras APA Implementadas

### Aspas Duplas ("  ")
✅ **Usar para:**
- Títulos de artigos, capítulos e webpages no corpo do texto
- Citações diretas curtas (< 40 palavras)
- Termos irônicos ou inventados (primeira ocorrência)
- Títulos de testes e escalas (quando mencionados no texto)

### Aspas Simples ('  ')
✅ **Usar para:**
- Citações dentro de citações
- Exemplo: "O autor afirmou que 'este é o melhor método' disponível" (Silva, 2020).

### Itálico (sem aspas)
✅ **Usar para:**
- Termos técnicos e estatísticos (ex: *p* < .05)
- Âncoras de escala (ex: *1 = discordo totalmente*)
- Títulos de livros, periódicos e filmes
- Introdução de novos termos técnicos

### Formato de Bloco (sem aspas)
✅ **Usar para:**
- Citações com 40 ou mais palavras
- Recuo de 1,27 cm da margem esquerda
- Sem aspas
- Referência após o ponto final

---

## 🚀 Instalação

### Pré-requisitos
- Microsoft Word 2016 ou superior
- Node.js 14+ (para desenvolvimento)
- npm 6+ (para desenvolvimento)

### Instalação do Add-in

#### Opção 1: Desenvolvimento Local

```bash
# Clone o repositório
git clone https://github.com/academic-tools/apa-formatter.git
cd apa-formatter

# Instale as dependências
npm install

# Gere certificados SSL para desenvolvimento
npm run dev-certs

# Compile o TypeScript
npm run build

# Inicie o servidor de desenvolvimento
npm start
```

#### Opção 2: Sideload Manual

1. Abra o Microsoft Word
2. Vá em **Inserir > Meus Suplementos > Gerenciar Meus Suplementos**
3. Clique em **Carregar Suplemento**
4. Selecione o arquivo `manifest.xml`
5. O add-in será carregado na guia "Página Inicial"

#### Opção 3: Implantação Corporativa

Para implantação em toda a organização, consulte:
[Documentação de Implantação Office](https://docs.microsoft.com/office/dev/add-ins/publish/publish)

---

## 💻 Como Usar

### Interface Principal

1. **Abra o Word** e carregue seu documento acadêmico
2. **Clique no botão "Formatar Aspas"** na guia Página Inicial
3. **Selecione as opções desejadas:**
   - ☑️ Converter Aspas Tipográficas
   - ☑️ Validar Uso de Aspas
   - ☑️ Corrigir Pontuação
   - ☑️ Identificar Citações Longas
   - ☑️ Gerar Relatório Detalhado
4. **Escolha o escopo:**
   - 🌐 Documento Inteiro
   - 📄 Apenas Seleção
5. **Clique em "Executar"**
6. **Revise o relatório** gerado automaticamente

### Formatação Rápida

Para formatação rápida com configurações padrão:
1. Clique no botão **"Formatação Rápida"** na Ribbon
2. As alterações serão aplicadas automaticamente ao documento inteiro

### Menu de Contexto

1. Selecione o texto que deseja formatar
2. Clique com o **botão direito**
3. Escolha **"Formatar Seleção (APA)"**

### Atalhos de Teclado

| Atalho | Ação |
|--------|------|
| `Ctrl + Alt + F` | Abrir painel de formatação |
| `Ctrl + Alt + Q` | Formatação rápida |

---

## 📖 Exemplos de Uso

### Exemplo 1: Citação Curta com Referência
**Antes:**
```
O autor afirmou que "este método é eficaz" (Silva, 2020, p. 15).
```

**Depois:**
```
O autor afirmou que "este método é eficaz" (Silva, 2020, p. 15).
```

### Exemplo 2: Citação Dentro de Citação
**Antes:**
```
"O pesquisador disse "aspas simples aqui" no estudo" (Silva, 2020).
```

**Depois:**
```
"O pesquisador disse 'aspas simples aqui' no estudo" (Silva, 2020).
```

### Exemplo 3: Citação Longa
**Antes:**
```
"Lorem ipsum dolor sit amet, consectetur adipiscing elit. Sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat." (Silva, 2020, p. 42).
```

**Depois (formato de bloco):**
```
    Lorem ipsum dolor sit amet, consectetur adipiscing elit. Sed do 
    eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim 
    ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut 
    aliquip ex ea commodo consequat. (Silva, 2020, p. 42)
```

### Exemplo 4: Termo Técnico
**Antes:**
```
O teste de "qui-quadrado" foi utilizado.
```

**Depois:**
```
O teste de qui-quadrado foi utilizado.
```
*Nota: Termos técnicos devem estar em itálico, não entre aspas.*

---

## 🛠️ Desenvolvimento

### Estrutura do Projeto

```
apa-quote-formatter/
├── apa-quote-formatter.ts    # Código principal TypeScript
├── taskpane.html              # Interface do usuário
├── commands.html              # Comandos executáveis
├── manifest.xml               # Manifesto do add-in
├── package.json               # Configuração npm
├── tsconfig.json              # Configuração TypeScript
├── README.md                  # Este arquivo
└── assets/                    # Ícones e recursos
    ├── icon-16.png
    ├── icon-32.png
    └── icon-80.png
```

### Compilar o Projeto

```bash
# Compilar uma vez
npm run build

# Compilar e observar mudanças
npm run watch
```

### Executar Testes

```bash
# Executar suite de testes
npm test

# Testes com cobertura
npm run test:coverage
```

### Validar Manifesto

```bash
npm run validate
```

---

## 🧪 Casos de Teste

O add-in foi testado extensivamente com os seguintes casos:

| Teste | Entrada | Saída Esperada | Status |
|-------|---------|----------------|--------|
| CT-01 | Aspas retas simples | Aspas tipográficas | ✅ Passou |
| CT-02 | Citação aninhada | Aspas duplas externas, simples internas | ✅ Passou |
| CT-03 | Citação 40+ palavras | Conversão para bloco | ✅ Passou |
| CT-04 | Termo técnico entre aspas | Alerta de validação | ✅ Passou |
| CT-05 | Pontuação incorreta | Correção automática | ✅ Passou |
| CT-06 | Âncora de escala | Alerta de validação | ✅ Passou |

---

## 🔧 Requisitos Técnicos

### Compatibilidade
- ✅ Microsoft Word 2016+
- ✅ Word Online
- ✅ Word para Mac
- ✅ Word para iPad

### Plataformas
- Windows 10+
- macOS 10.14+
- iOS 13+
- Web (navegadores modernos)

### Tecnologias
- TypeScript 5.0+
- Office.js API
- HTML5 + CSS3
- ES2020

---

## 🚨 Tratamento de Erros

O add-in implementa tratamento robusto de erros:

- ✅ Validação de entrada antes de processar
- ✅ Mensagens de erro descritivas
- ✅ Logging detalhado para debugging
- ✅ Rollback automático em caso de falha
- ✅ Opção de desfazer alterações (Ctrl+Z)

---

## 📚 Recursos e Referências

### Documentação APA
- [APA Style Official Website](https://apastyle.apa.org/)
- [APA Publication Manual 7th Edition](https://apastyle.apa.org/products/publication-manual-7th-edition)
- [APA Style Blog](https://apastyle.apa.org/blog)

### Documentação Office Add-ins
- [Office Add-ins Documentation](https://docs.microsoft.com/office/dev/add-ins/)
- [Word JavaScript API](https://docs.microsoft.com/javascript/api/word)
- [Office UI Fabric](https://developer.microsoft.com/fluentui)

---

## 🤝 Contribuindo

Contribuições são bem-vindas! Por favor:

1. Fork o projeto
2. Crie uma branch para sua feature (`git checkout -b feature/MinhaFeature`)
3. Commit suas mudanças (`git commit -m 'Adiciona MinhaFeature'`)
4. Push para a branch (`git push origin feature/MinhaFeature`)
5. Abra um Pull Request

### Diretrizes de Contribuição
- Siga o estilo de código TypeScript estabelecido
- Adicione testes para novas funcionalidades
- Atualize a documentação conforme necessário
- Mantenha commits atômicos e descritivos

---

## 📄 Licença

Este projeto está licenciado sob a Licença MIT - veja o arquivo [LICENSE](LICENSE) para detalhes.

---

## 👨‍💻 Autor

**Academic Tools Team**
- Website: [https://academic-tools.github.io](https://academic-tools.github.io)
- Email: contact@academic-tools.dev
- GitHub: [@academic-tools](https://github.com/academic-tools)

---

## 🙏 Agradecimentos

- Equipe da Microsoft Office por fornecer excelente documentação
- Comunidade APA Style por manter padrões claros
- Todos os contribuidores e testadores beta

---

## 📞 Suporte

Encontrou um bug ou tem uma sugestão?

- 🐛 [Reportar Bug](https://github.com/academic-tools/apa-formatter/issues/new?template=bug_report.md)
- 💡 [Sugerir Feature](https://github.com/academic-tools/apa-formatter/issues/new?template=feature_request.md)
- 📧 Email: support@academic-tools.dev
- 💬 [Discussões](https://github.com/academic-tools/apa-formatter/discussions)

---

## 📊 Changelog

### [1.0.0] - 2024-02-13
- 🎉 Lançamento inicial
- ✨ Conversão de aspas retas para tipográficas
- ✨ Validação segundo normas APA 7ª edição
- ✨ Correção automática de pontuação
- ✨ Identificação de citações longas
- ✨ Relatório detalhado de alterações
- ✨ Interface responsiva e moderna

---

## 🔮 Roadmap

### Versão 1.1.0 (Planejado)
- [ ] Suporte para formatação de referências bibliográficas
- [ ] Integração com Zotero e Mendeley
- [ ] Verificação de formato de números e datas
- [ ] Modo de sugestão (track changes)

### Versão 1.2.0 (Futuro)
- [ ] Suporte multilíngue (inglês, espanhol)
- [ ] Templates de documentos APA
- [ ] Verificação de plágio básica
- [ ] Exportação de relatórios em PDF

---

## ⚖️ Aviso Legal

Este add-in é uma ferramenta de auxílio à formatação e não substitui a revisão manual e o conhecimento das normas APA. Sempre revise seu documento final para garantir conformidade total com os padrões acadêmicos de sua instituição.

O autor não se responsabiliza por erros ou inconsistências resultantes do uso desta ferramenta. Use por sua conta e risco.

---

**Desenvolvido com ❤️ para a comunidade acadêmica**

*Última atualização: 13 de fevereiro de 2024*
