import smtplib
import time
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

def enviar_emails():
    print("="*50)
    print("🚀 ENVIO DE E-MAILS AUTOMÁGICO (SCRIPT LOCAL)")
    print("="*50)
    
    # Configurações do email
    email_remetente = input("👉 O seu E-mail da Empresa (Gmail): ").strip()
    password_app = input("👉 A sua 'App Password' (Palavra-passe de Aplicação): ").strip()
    assunto = input("👉 Assunto do E-mail: ").strip()
    
    print("\n👉 Escreva o corpo do e-mail (escreva 'FIM' numa nova linha para terminar a mensagem):")
    linhas_corpo = []
    while True:
        linha = input()
        if linha.strip().upper() == 'FIM':
            break
        linhas_corpo.append(linha)
    corpo_mensagem = "\n".join(linhas_corpo)
    
    # LER DESTINATÁRIOS
    # O script vai ler o ficheiro destinatarios.txt que tem de estar na mesma pasta
    arquivos_emails = "destinatarios.txt"
    try:
        with open(arquivos_emails, "r", encoding="utf-8") as file:
            destinatarios = [linha.strip() for linha in file if "@" in linha] # Simples validação se tem @
    except FileNotFoundError:
        print(f"\n❌ ERRO: Ficheiro '{arquivos_emails}' não encontrado.")
        print("Crie um ficheiro de texto chamado 'destinatarios.txt' com um email em cada linha e volte a correr o script.")
        return

    if not destinatarios:
        print("\n❌ ERRO: O ficheiro 'destinatarios.txt' está vazio ou não tem e-mails válidos.")
        return

    print(f"\n🕒 Lidos {len(destinatarios)} destinatários. A ligar ao servidor da Google...")

    servidor_smtp = "smtp.gmail.com"
    porta = 587

    try:
        # Ligar ao servidor SMTP do Google
        server = smtplib.SMTP(servidor_smtp, porta)
        server.starttls() # Securizar a ligação
        server.login(email_remetente, password_app)
        
        print("✅ Ligação efetuada com sucesso! A iniciar envio...\n")
        
        sucessos = 0
        erros = 0

        # Enviar um e-mail SEPARADO para cada colaborador (Mail Merge básico)
        for destinatario in destinatarios:
            try:
                msg = MIMEMultipart()
                msg['From'] = email_remetente
                msg['To'] = destinatario
                msg['Subject'] = assunto
                
                # Adicionar o corpo à estrutura do email
                msg.attach(MIMEText(corpo_mensagem, 'plain', 'utf-8'))
                
                # Executar envio
                server.send_message(msg)
                print(f"   📧 ✅ Enviado: {destinatario}")
                sucessos += 1
                
                # Pequena pausa de 1.5s para não ativar os limites de 'Spam' da Google
                time.sleep(1.5)
                
            except Exception as e:
                print(f"   📧 ❌ Falhou ({destinatario}): {e}")
                erros += 1

        # Fechar ligação
        server.quit()
        print("\n" + "="*50)
        print(f"🎉 CONCLUÍDO: {sucessos} e-mails enviados. {erros} erros.")
        print("="*50)

    except smtplib.SMTPAuthenticationError:
        print("\n❌ ERRO DE AUTENTICAÇÃO: A Google bloqueou o acesso.")
        print("Certifique-se que o email está correto e que gerou uma 'Password de Aplicação' (App Password) na sua conta Google.")
        print("A sua password normal do Gmail não funciona para este tipo de scripts de segurança.")
    except Exception as e:
        print(f"\n❌ ERRO FATAL: {e}")

if __name__ == "__main__":
    enviar_emails()
