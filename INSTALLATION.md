# 📦 Guia de Instalação - Formatador APA

## Guia Completo para Instalação e Configuração

Este guia fornece instruções detalhadas para instalar e configurar o **Formatador APA** no Microsoft Word.

---

## 🎯 Pré-requisitos

Antes de começar, certifique-se de ter:

### Software Necessário
- ✅ **Microsoft Word 2016 ou superior** (Windows, Mac ou Online)
- ✅ **Node.js 14+** (apenas para desenvolvimento)
- ✅ **npm 6+** (instalado automaticamente com Node.js)
- ✅ **Git** (para clonar o repositório)

### Conhecimentos Recomendados
- 📚 Noções básicas de linha de comando
- 📚 Familiaridade com o Microsoft Word
- 📚 (Opcional) Conhecimento básico de TypeScript/JavaScript

---

## 📥 Opção 1: Instalação para Usuários (Simples)

### Passo 1: Baixar os Arquivos

1. Acesse o [repositório no GitHub](https://github.com/academic-tools/apa-formatter)
2. Clique no botão verde **"Code"**
3. Selecione **"Download ZIP"**
4. Extraia o arquivo ZIP em uma pasta de sua preferência

### Passo 2: Carregar no Word (Sideload)

#### Windows:

1. Abra o **Microsoft Word**
2. Vá em **Arquivo > Opções > Central de Confiabilidade**
3. Clique em **Configurações da Central de Confiabilidade**
4. Selecione **Catálogos de Suplementos Confiáveis**
5. Adicione a pasta onde você extraiu os arquivos
6. Marque **"Mostrar no menu"**
7. Clique em **OK** e reinicie o Word

#### Mac:

1. Abra o **Microsoft Word**
2. Vá em **Inserir > Suplementos > Meus Suplementos**
3. Clique em **"+"** (Adicionar suplemento)
4. Selecione **"Adicionar de um arquivo"**
5. Navegue até a pasta extraída
6. Selecione o arquivo **manifest.xml**
7. Clique em **OK**

#### Word Online:

1. Acesse [Office.com](https://office.com) e abra um documento
2. Clique em **Inserir > Suplementos do Office**
3. Selecione **Carregar Meu Suplemento**
4. Clique em **Procurar** e selecione **manifest.xml**
5. Clique em **Carregar**

### Passo 3: Ativar o Add-in

1. Na guia **Página Inicial**, você verá o novo grupo **"Formatador APA"**
2. Clique no botão **"Formatar Aspas"** para abrir o painel
3. Pronto! O add-in está instalado e pronto para uso

---

## 🛠️ Opção 2: Instalação para Desenvolvedores (Avançada)

### Passo 1: Configurar Ambiente de Desenvolvimento

#### 1.1 Instalar Node.js

**Windows:**
```bash
# Baixar do site oficial
https://nodejs.org/

# Ou usar Chocolatey
choco install nodejs
```

**Mac:**
```bash
# Usando Homebrew
brew install node
```

**Linux:**
```bash
# Ubuntu/Debian
sudo apt update
sudo apt install nodejs npm

# Fedora
sudo dnf install nodejs
```

#### 1.2 Verificar Instalação

```bash
node --version  # Deve mostrar v14.0.0 ou superior
npm --version   # Deve mostrar 6.0.0 ou superior
```

### Passo 2: Clonar e Configurar Projeto

```bash
# Clonar repositório
git clone https://github.com/academic-tools/apa-formatter.git
cd apa-formatter

# Instalar dependências
npm install

# Isso pode levar alguns minutos...
```

### Passo 3: Gerar Certificados SSL

Para desenvolvimento local, você precisa de certificados SSL:

```bash
# Instalar ferramenta de certificados
npm install -g office-addin-dev-certs

# Gerar certificados
npx office-addin-dev-certs install

# Confiar nos certificados (vai pedir senha de administrador)
```

**Nota:** No Windows, pode aparecer um aviso de segurança. Clique em **"Sim"** para confiar.

### Passo 4: Compilar TypeScript

```bash
# Compilar o código TypeScript
npm run build

# Você verá uma pasta 'dist' sendo criada
```

### Passo 5: Iniciar Servidor de Desenvolvimento

```bash
# Iniciar servidor local
npm start

# O servidor iniciará em https://localhost:3000
```

**Importante:** Mantenha esta janela do terminal aberta enquanto desenvolve!

### Passo 6: Carregar no Word (Desenvolvimento)

#### Windows:

```bash
# Abrir Word e carregar automaticamente
npm run sideload

# Ou manualmente:
# 1. Abra Word
# 2. Insira > Meus Suplementos > Gerenciar Meus Suplementos
# 3. Carregar Suplemento > Procurar
# 4. Selecione manifest.xml da pasta do projeto
```

#### Mac:

```bash
# Comando para sideload no Mac
npm run sideload:mac

# Ou manualmente seguindo as instruções acima
```

### Passo 7: Desenvolvimento com Hot Reload

O servidor de desenvolvimento suporta hot reload:

1. Faça mudanças no código TypeScript
2. O código será automaticamente recompilado
3. Recarregue o painel do add-in no Word (F5)

---

## 🔧 Configuração Avançada

### Configurar Porta Customizada

Edite o arquivo `package.json`:

```json
{
  "scripts": {
    "start": "office-addin-dev-server --dev-server-port 8080"
  }
}
```

Depois atualize o `manifest.xml` com a nova porta.

### Configurar Debugging

#### Visual Studio Code:

1. Instale a extensão **"Office Add-in Debugger"**
2. Pressione **F5** para iniciar debugging
3. Selecione **"Word Desktop"** ou **"Word Online"**

Configuração `.vscode/launch.json`:

```json
{
  "version": "0.2.0",
  "configurations": [
    {
      "name": "Office Add-in (Desktop)",
      "type": "office-addin",
      "request": "attach",
      "port": 9229,
      "timeout": 60000
    }
  ]
}
```

#### Chrome DevTools (Word Online):

1. Abra Word Online com o add-in carregado
2. Pressione **F12** para abrir DevTools
3. Vá em **Console** para ver logs
4. Use **Debugger** para breakpoints

---

## 📱 Instalação em Plataformas Específicas

### Word para iPad

1. Instale o app **Microsoft Word** da App Store
2. Abra um documento
3. Toque no ícone **"+"** no canto superior direito
4. Selecione **Suplementos > Gerenciar Suplementos**
5. Toque em **Carregar Suplemento**
6. Selecione o arquivo manifest.xml (via iCloud Drive ou OneDrive)

### Word Mobile (Android)

1. Instale **Microsoft Word** da Play Store
2. Abra um documento
3. Toque nos **três pontos** (⋮) no canto superior direito
4. Selecione **Suplementos**
5. Toque em **Mais Suplementos**
6. Escolha **Carregar Suplemento Personalizado**

---

## 🚨 Solução de Problemas

### Problema: "Não consigo ver o add-in no Word"

**Solução:**
1. Verifique se o manifest.xml está válido:
   ```bash
   npm run validate
   ```
2. Reinicie o Word completamente
3. Verifique se a pasta está em Catálogos Confiáveis (Windows)

### Problema: "Erro de certificado SSL"

**Solução:**
```bash
# Reinstalar certificados
npx office-addin-dev-certs install --force

# No Mac, pode precisar de:
sudo security add-trusted-cert -d -r trustRoot -k ~/Library/Keychains/login.keychain-db ~/.office-addin-dev-certs/ca.crt
```

### Problema: "Add-in não carrega / tela branca"

**Solução:**
1. Abra Console do desenvolvedor (F12)
2. Verifique erros no console
3. Certifique-se de que o servidor está rodando:
   ```bash
   npm start
   ```
4. Limpe o cache do Office:
   - Windows: `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef`
   - Mac: `~/Library/Containers/com.microsoft.Word/Data/Library/Application Support/Microsoft/Office/16.0/Wef/`

### Problema: "npm install falha"

**Solução:**
```bash
# Limpar cache npm
npm cache clean --force

# Remover node_modules
rm -rf node_modules package-lock.json

# Reinstalar
npm install
```

### Problema: "TypeScript não compila"

**Solução:**
1. Verifique a versão do TypeScript:
   ```bash
   npm list typescript
   ```
2. Se necessário, reinstale:
   ```bash
   npm install typescript@latest --save-dev
   ```

---

## 📝 Checklist de Instalação

Use este checklist para garantir que tudo está funcionando:

- [ ] Node.js instalado (v14+)
- [ ] npm instalado (v6+)
- [ ] Repositório clonado
- [ ] Dependências instaladas (`npm install`)
- [ ] Certificados SSL gerados
- [ ] TypeScript compilado (`npm run build`)
- [ ] Servidor iniciado (`npm start`)
- [ ] manifest.xml validado (`npm run validate`)
- [ ] Add-in carregado no Word
- [ ] Painel do add-in abre corretamente
- [ ] Funções básicas testadas

---

## 🔄 Atualização

Para atualizar para a versão mais recente:

```bash
# Navegar até a pasta do projeto
cd apa-formatter

# Obter atualizações
git pull origin main

# Reinstalar dependências (se necessário)
npm install

# Recompilar
npm run build

# Reiniciar servidor
npm start
```

---

## 🗑️ Desinstalação

### Remover do Word:

1. Abra Word
2. Vá em **Inserir > Meus Suplementos**
3. Clique em **Gerenciar Meus Suplementos**
4. Encontre **Formatador APA**
5. Clique no **ícone de lixeira** ou **Remover**

### Remover Arquivos do Projeto:

```bash
# Apenas remover a pasta
rm -rf apa-formatter

# Ou no Windows
rmdir /s apa-formatter
```

### Remover Certificados SSL:

```bash
npx office-addin-dev-certs uninstall
```

---

## 📞 Suporte para Instalação

Problemas durante a instalação?

- 📧 Email: support@academic-tools.dev
- 💬 [GitHub Issues](https://github.com/academic-tools/apa-formatter/issues)
- 📚 [Documentação Oficial](https://docs.microsoft.com/office/dev/add-ins/)
- 🎥 [Vídeo Tutorial](https://youtube.com/watch?v=example) (em breve)

---

## 🎓 Recursos Adicionais

- [Documentação Office Add-ins](https://docs.microsoft.com/office/dev/add-ins/)
- [Yeoman Generator para Office](https://github.com/OfficeDev/generator-office)
- [Script Lab (para testar APIs)](https://aka.ms/getscriptlab)
- [Office Dev Center](https://developer.microsoft.com/office)

---

**Última atualização:** 13 de fevereiro de 2024

**Próximos passos:** Após a instalação, consulte o [Guia de Uso](USAGE_GUIDE.md) para aprender a usar todas as funcionalidades.
