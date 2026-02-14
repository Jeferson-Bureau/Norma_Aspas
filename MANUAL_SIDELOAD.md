# Sideload Manual do Add-in

Se o carregamento automático falhar, siga estes passos para carregar o add-in manualmente no Word.

## Opção 1: Usando "Meus Suplementos" (Mais fácil)

1. Abra o **Microsoft Word**.
2. Crie um documento em branco.
3. Vá na guia **Inserir** > **Meus Suplementos** (ou "Suplementos").
4. Procure por uma guia ou link chamado **"Gerenciar Meus Suplementos"** ou **"Carregar Suplemento"** (Upload My Add-in).
    * *Nota: Esta opção pode não aparecer em todas as versões desktop, sendo mais comum no Word Online ou Mac.*
5. Se disponível, selecione **"Carregar Suplemento"** e escolha o arquivo `manifest.xml` na pasta do projeto:
    `D:\04_UEM\Revistas\ACTA_SCIENTIARUM\Script\Norma_Aspas\manifest.xml`

## Opção 2: Catálogo de Pasta Compartilhada (Padrão Windows)

Se a opção acima não existir:

1. **Crie uma pasta** no seu computador (ex: `C:\MyAddins`).
2. **Copie** o arquivo `manifest.xml` do projeto para esta pasta.
3. Abra o Word e vá em **Arquivo > Opções**.
4. Clique em **Central de Confiabilidade** > **Configurações da Central de Confiabilidade**.
5. Selecione **Catálogos de Suplementos Confiáveis** no menu esquerdo.
6. Em "URL do Catálogo", digite o caminho da pasta: `C:\MyAddins`.
7. Clique em **Adicionar Catálogo**.
8. Marque a caixa **"Mostrar no Menu"**.
9. Clique em **OK** e reinicie o Word.
10. Agora, vá em **Inserir > Meus Suplementos > Pasta Compartilhada**, e o "Formatador APA" deve aparecer lá.

## Verificação do Servidor

Para que o add-in funcione, o "servidor" (janela preta do terminal) deve permanecer aberto.
Se fechou, execute na pasta do projeto:

```bash
npm run dev-server
```

## Solução de Problemas Comuns

### Erro: "Este suplemento não está mais disponível"

1. **Verifique se o servidor está rodando:**
    Abra seu navegador (Edge ou Chrome) e acesse:
    [https://localhost:3000/taskpane.html](https://localhost:3000/taskpane.html)

    * **Se abrir uma página em branco ou com erro de certificado:** Prossiga para o passo 2.
    * **Se não carregar nada:** O servidor não está rodando. Volte ao terminal e execute `npm start`.

2. **Certificados de Segurança:**
    Se o navegador mostrar "Sua conexão não é privada", o Word bloqueará o add-in.
    Execute no terminal para corrigir:

    ```bash
    npm run dev-certs
    ```

3. **Limpar Cache do Office:**
    Às vezes o Word "trava" em uma versão antiga.
    * Feche o Word.
    * Exclua a pasta: `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\`
    * Reinicie o `npm start`.
