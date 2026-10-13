# Email Flash ⚡

> Desenvolvido por mim com ajuda de IA (Claude) / Built by me with AI help (Claude)

**PT:** Aplicativo desktop em Python (Tkinter) que automatiza o envio de notas fiscais (DANFE) e boletos por e-mail, anexando os PDFs certos e salvando uma cópia na pasta Enviados.

**EN:** Python (Tkinter) desktop app that automates sending invoices (DANFE) and payment slips (boletos) by email, attaching the right PDFs and saving a copy in the Sent folder.

## Por que existe / Why it exists
Criei para o meu trabalho no faturamento de uma distribuidora, onde enviar NFs e boletos um a um tomava muito tempo. Uso no dia a dia.

I built it for my billing job at a distribution company, where sending invoices and boletos one by one took a lot of time. I use it daily.

## Funcionalidades / Features
- Envio manual: você digita o número e a série da NF (ex.: `99999-1`) / Manual sending: type the invoice number and series (e.g. `99999-1`)
- Várias NFs no mesmo e-mail, separadas por vírgula / Multiple invoices in one email, separated by commas
- Campo separado para NFs sem boleto (bonificação, à vista) / Separate field for invoices without a boleto
- Envio em lote a partir de uma planilha Excel (.xls/.xlsx) / Batch sending from an Excel spreadsheet (.xls/.xlsx)
- Localiza sozinho os PDFs da NF e do boleto nas pastas configuradas / Finds the invoice and boleto PDFs automatically in the configured folders
- Assunto e corpo do e-mail personalizáveis com variáveis / Customizable subject and body with variables
- Campos CC e CCO / CC and BCC fields
- Assinatura em imagem no final do e-mail / Image signature at the end of the email
- Conexão SSL ou STARTTLS / SSL or STARTTLS connection
- Salva uma cópia do e-mail enviado na pasta Enviados via IMAP / Saves a copy of the sent email in the Sent folder via IMAP
- Senha guardada com segurança no cofre do sistema (keyring) / Password stored securely in the system vault (keyring)
- Interface que não trava durante o envio (threads) / UI stays responsive while sending (threads)

## Tecnologias / Tech
Python, Tkinter, smtplib, imaplib, keyring, openpyxl, xlrd

## Como usar / How to use
1. Instale o Python e as bibliotecas / Install Python and the libraries:
   `pip install -r requirements.txt`
2. Execute / Run: `python Email_Flash.py`
3. Na aba **Config. E-mail**, informe seu e-mail, senha e os servidores SMTP e IMAP.
   In the **Config. E-mail** tab, enter your email, password, and the SMTP and IMAP servers.
4. Na aba **Config. Gerais**, escolha a pasta das NFs, a pasta dos boletos e a imagem de assinatura.
   In the **Config. Gerais** tab, choose the invoices folder, the boletos folder, and the signature image.
5. Na aba **Envio**, digite a NF, o nome e o e-mail do cliente e clique em **Enviar E-mail**.
   In the **Envio** tab, type the invoice, customer name and email, then click **Enviar E-mail**.

**Nomes de arquivo esperados / Expected file names:**
- NF / Invoice: `danfe_SERIE_NUMERO.pdf`
- Boleto: `boleto nf 00000000.pdf` (número com 8 dígitos / 8-digit number)

**Envio em lote / Batch sending:** a primeira linha da planilha deve ter as colunas `SERIE NF`, `NUMERO NF`, `CLIENTE` e `EMAIL`. / The first row must have the columns `SERIE NF`, `NUMERO NF`, `CLIENTE` and `EMAIL`.

## Privacidade / Privacy
Suas configurações ficam no seu computador, em `%APPDATA%\EmailFlash`, e a senha fica no cofre do sistema. Nada é enviado para este repositório.

Your settings stay on your computer, in `%APPDATA%\EmailFlash`, and the password is stored in the system vault. Nothing is sent to this repository.

## Sobre o uso de IA / About AI use
Este projeto foi desenvolvido por mim com ajuda de inteligência artificial (Claude). Eu defini o que o programa precisa fazer, testei e uso no trabalho; a IA ajudou a escrever e melhorar o código.

I developed this project with the help of AI (Claude). I defined what the program needs to do, tested it, and use it at work; the AI helped write and improve the code.

## Autor / Author
Richardson — [LinkedIn](https://www.linkedin.com/in/richardsongp)
