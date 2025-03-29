import smtplib
import os
import csv
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
from email.mime.base import MIMEBase
from email import encoders
import time  # <<< 1. Importar o módulo time

# --- Início da contagem de tempo ---
start_time = time.time()  # <<< 2. Guardar o tempo de início

# Configurações do e-mail e servidor SMTP
EMAIL = "seu@email.com"
SMTP_SERVER = "smtp.gmail.com"
SMTP_PORT = 465

# Pasta onde estão os PDFs
PASTA_PDFS = "C:\\pc\\user\\arquivos_pdfs"

# Lendo a senha do arquivo
try:
    with open('senha.txt') as f:
        senha_do_email = f.readline().strip()
except FileNotFoundError:
    print("❌ Erro Crítico: Arquivo 'capital-social-senha.txt' não encontrado.")
    exit() # Sai do script se não encontrar a senha
except Exception as e:
    print(f"❌ Erro Crítico ao ler o arquivo de senha: {e}")
    exit()

# Listas para armazenar os e-mails enviados com sucesso e os não encontrados
enviados = []
nao_encontrados = []

# Lendo os nomes e e-mails do arquivo CSV
try:
    with open('lista_de_email.csv', newline='', encoding='utf-8') as csvfile:
        leitor_csv = csv.reader(csvfile)
        try:
            next(leitor_csv)  # Pula o cabeçalho, se houver
        except StopIteration:
            print("⚠️ Arquivo CSV 'capital-social.csv' está vazio ou não tem cabeçalho.")
            # Considerar se deve continuar ou sair
            pass # Continua mesmo sem cabeçalho ou vazio

        print("\n--- Iniciando processamento dos e-mails ---")
        for i, linha in enumerate(leitor_csv, start=1): # Adiciona contador de linha
            if len(linha) < 2:
                print(f"⚠️ Linha {i}: Incompleta encontrada: {linha}. Pulando esta linha.")
                continue
            nome = linha[0].strip()
            destinatario = linha[1].strip()

            # Verifica se nome ou destinatario estão vazios
            if not nome or not destinatario:
                print(f"⚠️ Linha {i}: Nome ou destinatário vazio. Nome='{nome}', Email='{destinatario}'. Pulando esta linha.")
                continue

            pdf_filename = f"{nome}.pdf"
            pdf_path = os.path.join(PASTA_PDFS, pdf_filename)

            # Verifica se o arquivo PDF existe para o destinatário
            if not os.path.exists(pdf_path):
                print(f"⚠️ Linha {i}: Nenhum PDF encontrado para '{nome}' ({destinatario}) em '{pdf_path}'. E-mail NÃO enviado.")
                nao_encontrados.append([nome, destinatario, "PDF não encontrado"]) # <<< CORREÇÃO: Adiciona à lista nao_encontrados
                continue # Pula o envio caso não ache o PDF.

            # --- Construção do Email ---
            msg = MIMEMultipart("related")
            msg['Subject'] = 'Demonstrativo Capital Social 2024 Unimed Gerais de Minas'
            msg['From'] = EMAIL
            msg['To'] = destinatario

            html = f"""
            <html>
                <body>
                    <p><strong>Prezado(a) {nome},</strong></p></br>
                    </br>
                    <p>Em anexo, você encontrará o seu demonstrativo do capital social referente ao ano de 2024.📊</p>
                    <p>Caso tenha alguma dúvida ou precise de esclarecimento adicionais, por favor, entre em contato conosco pelo telefone (38) 3729-7249 e fale com a Grace. 📞</p>
                    <p>Atenciosamente,</p>
                    <div style="text-align: center;">
                    <p>
                    <img src="cid:image.png" alt="Mural" style="width:100%; max-width:400px;" />
                    </div>
                </body>
            </html>
            """
            msg.attach(MIMEText(html, "html"))

            # Embutir a imagem (se existir)
            try:
                with open("image.png", "rb") as img_file:
                    img = MIMEImage(img_file.read())
                    img.add_header("Content-ID", "<image.png>")
                    img.add_header("Content-Disposition", "inline", filename="image.png")
                    msg.attach(img)
            except FileNotFoundError:
                print(f"⚠️ Linha {i}: Arquivo 'image.png' não encontrado. Imagem não será embutida para '{nome}'.")
            except Exception as e:
                 print(f"⚠️ Linha {i}: Erro ao anexar imagem para '{nome}': {e}")

            # Anexando o PDF
            try:
                with open(pdf_path, "rb") as pdf_file:
                    part = MIMEBase("application", "octet-stream")
                    part.set_payload(pdf_file.read())
                    encoders.encode_base64(part)
                    part.add_header("Content-Disposition", f"attachment; filename=\"{pdf_filename}\"") # Adiciona aspas ao nome do arquivo
                    msg.attach(part)
            except Exception as e:
                print(f"❌ Linha {i}: Erro ao anexar PDF para '{nome}' ({destinatario}): {e}")
                nao_encontrados.append([nome, destinatario, f"Erro ao anexar PDF: {e}"]) # Adiciona à lista nao_encontrados
                continue # Pula o envio se não conseguir anexar

            # --- Enviando o e-mail ---
            try:
                # Tenta conectar e enviar
                with smtplib.SMTP_SSL(SMTP_SERVER, SMTP_PORT, timeout=60) as smtp: # Adiciona timeout
                    smtp.login(EMAIL, senha_do_email)
                    smtp.send_message(msg)
                print(f"✅ Linha {i}: E-mail enviado para '{nome}' ({destinatario})")
                enviados.append([nome, destinatario])
            except smtplib.SMTPAuthenticationError:
                 print(f"❌ Linha {i}: Erro de autenticação ao enviar para '{nome}' ({destinatario}). Verifique e-mail/senha.")
                 nao_encontrados.append([nome, destinatario, "Erro de autenticação SMTP"])
            except smtplib.SMTPServerDisconnected:
                 print(f"❌ Linha {i}: Servidor SMTP desconectou ao enviar para '{nome}' ({destinatario}).")
                 nao_encontrados.append([nome, destinatario, "Servidor SMTP desconectado"])
            except Exception as e:
                print(f"❌ Linha {i}: Erro ao enviar e-mail para '{nome}' ({destinatario}): {e}")
                nao_encontrados.append([nome, destinatario, f"Erro envio: {e}"])

            # Pausa opcional para não sobrecarregar o servidor ou ser bloqueado
            # time.sleep(1) # Pausa por 1 segundo entre os envios

except FileNotFoundError:
    print(f"❌ Erro Crítico: Arquivo CSV 'capital-social.csv' não encontrado.")
    exit()
except Exception as e:
    print(f"❌ Erro Crítico ao processar o arquivo CSV: {e}")
    exit()

print("\n--- Finalizando registros ---")

# Salvar os e-mails enviados em um arquivo CSV
# (Removido o bloco duplicado que existia antes)
if enviados:
    try:
        # Usar 'w' para criar/sobrescrever ou 'a' para adicionar ao final
        with open('enviados.csv', mode='w', newline='', encoding='utf-8') as csvfile:
            writer = csv.writer(csvfile)
            writer.writerow(['Nome', 'Email']) # Adiciona cabeçalho
            writer.writerows(enviados)
        print(f"✅ Lista de {len(enviados)} e-mails enviados registrada em 'enviados.csv'.")
    except Exception as e:
        print(f"❌ Erro ao salvar 'enviados.csv': {e}")
else:
    print("⚠️ Nenhum e-mail foi enviado com sucesso.")

# Salvar os e-mails não encontrados em um arquivo CSV
if nao_encontrados:
    try:
        with open('nao_encontrados.csv', mode='w', newline='', encoding='utf-8') as csvfile:
            writer = csv.writer(csvfile)
            writer.writerow(['Nome', 'Email', 'Motivo']) # Adiciona cabeçalho
            writer.writerows(nao_encontrados)
        print(f"✅ Lista de {len(nao_encontrados)} e-mails NÃO enviados registrada em 'nao_encontrados.csv'.")
    except Exception as e:
         print(f"❌ Erro ao salvar 'nao_encontrados.csv': {e}")
else:
    # Se nao_encontrados estiver vazia E enviados também, significa que o CSV inicial pode estar vazio
    if not enviados:
         print("ℹ️ Nenhuma informação processada (CSV inicial pode estar vazio ou todas as linhas foram puladas).")
    else:
         print("✅ Todos os e-mails processados foram enviados com sucesso (nenhum PDF não encontrado ou erro de envio).")


# --- Fim da contagem de tempo ---
end_time = time.time()  # <<< 3. Guardar o tempo de fim
duration_seconds = end_time - start_time  # <<< 4. Calcular a duração

# --- Formatar e Imprimir o tempo total ---
minutes = int(duration_seconds // 60)
seconds = duration_seconds % 60
print("\n--- Script Finalizado ---")
print(f"Tempo total de execução: {minutes} minutos e {seconds:.2f} segundos.") # <<< 5. Imprimir resultado