# Script de Email 

Este script automatiza o envio de emails para os seus contatos.

**Objetivo:** Enviar um email personalizado.

**Pré-requisitos:**

*   `smtplib`:  Módulo Python para enviar e-mails.
*   Arquivo `senha.txt` (opcional, mas recomendado):  Seu login para o servidor SMTP.

**Como usar:**

1.  **Salve o script:** Salve o código como um arquivo `.py` (por exemplo, `bot-email.py`).
2.  **Execute o script:** Execute o script no terminal: `python bot-email.py`
3.  **Verifique a saída:** O script imprimirá informações sobre o processo de envio e o tempo de execução.

**Detalhes do Script:**

*   **Autenticação:** O script usa uma senha (armazenada em `senha.txt`) para se conectar ao servidor SMTP.
     Você pode usar variáveis de ambiente ou um arquivo de configuração mais seguro para armazenar a senha.
*   **Enviando o Email:** O script envia um email personalizado.
*   **Gerenciamento de Erros:** Implementação de tratamento básico de erros, como falhas na conexão e autenticação.

**Configurações:**

*   **SMTP Server:** `smtp.gmail.com` (ou outro servidor SMTP que você usa)
*   **Port:** 465 (padrão para Gmail)
*   **Senha:** `senha.txt` (ou arquivo de configuração seguro)

**Conteúdo do Email:**

*   Assunto: Demonstração 
*   Corpo:  (Inclui um texto de exemplo com os seus dados e o demonstrativo)
*   Imagem:  [Link para a imagem do demonstrativo]

**Importante:**

*   Certifique-se que você tem permissões de acesso ao arquivo `senha.txt` (ou equivalente).
*   Se usar uma senha, certifique-se de que o servidor SMTP esteja configurado corretamente.
*   Considere usar variáveis de ambiente para armazenar a senha e facilitar a configuração.

**Melhorias e Explicações:**

1. Tratamento de Exceções Melhorado: Adicionei blocos try...except mais amplos para capturar erros comuns, como:
    * FileNotFoundError ao ler o arquivo de senha.
    * smtplib.SMTPAuthenticationError: Para problemas de autenticação (verificar e-mail/senha).
    * smtplib.SMTPServerDisconnected: Se o servidor SMTP cair.
    * Erros durante a anexação do PDF.
      
2.**Lógica de Envio Redundante:** Adicionei try...except dentro da seção de envio do e-mail para lidar com possíveis erros durante a conexão, login ou envio do e-mail.
3. **Tratamento de Erro de Arquivo CSV:** Verificando se o arquivo CSV existe antes de continuar o processamento.
4. **Validação de Dados:** Adicionei verificações adicionais para garantir que os nomes e endereços de e-mail estejam preenchidos com informações válidas.
5. **Comentários Mais Detalhados:** Adicionei comentários mais explicativos em cada etapa do código.
6. **Formatação da Saída:** Formatei a saída para exibir o tempo total de execução de forma clara, usando f-strings e :.2f para formatar os segundos.
7. **Legibilidade:** Melhorei a formatação e indentação do código.
