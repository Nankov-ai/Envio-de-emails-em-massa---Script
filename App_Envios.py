import smtplib
import time
import os
import csv
import threading
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication

class EmailSenderApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Portal de Envio de E-mails - Norauto")
        self.root.geometry("650x750")
        self.root.resizable(False, False)
        
        # Variáveis
        self.caminho_ficheiro = tk.StringVar()
        self.caminho_anexos = tk.StringVar()
        self.emailVar = tk.StringVar()
        self.passVar = tk.StringVar()
        self.assuntoVar = tk.StringVar()
        self.destinatarios = [] # Lista de dicionários: [{'Nome': 'Joao', 'Email': 'joao@...'}, ...]
        
        self.construir_interface()

    def construir_interface(self):
        style = ttk.Style()
        style.theme_use('clam')
        
        main_frame = ttk.Frame(self.root, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)

        lbl_titulo = ttk.Label(main_frame, text="Envios Diretos (Personalizados + Anexos)", font=("Segoe UI", 16, "bold"))
        lbl_titulo.grid(row=0, column=0, columnspan=2, pady=(0, 20), sticky="w")

        # 1. Credenciais
        lbl_credenciais = ttk.Label(main_frame, text="1. Credenciais Google", font=("Segoe UI", 11, "bold"))
        lbl_credenciais.grid(row=1, column=0, columnspan=2, sticky="w", pady=(5, 5))

        ttk.Label(main_frame, text="O seu E-mail (Gmail):").grid(row=2, column=0, sticky="w", pady=2)
        ttk.Entry(main_frame, textvariable=self.emailVar, width=40).grid(row=2, column=1, sticky="w", pady=2)

        ttk.Label(main_frame, text="'App Password' (Google):").grid(row=3, column=0, sticky="w", pady=2)
        ttk.Entry(main_frame, textvariable=self.passVar, show="*", width=40).grid(row=3, column=1, sticky="w", pady=2)

        # 2. Ficheiros e Pastas
        lbl_ficheiro = ttk.Label(main_frame, text="2. Dados e Anexos", font=("Segoe UI", 11, "bold"))
        lbl_ficheiro.grid(row=4, column=0, columnspan=2, sticky="w", pady=(15, 5))

        frame_file = ttk.Frame(main_frame)
        frame_file.grid(row=5, column=0, columnspan=2, sticky="we", pady=5)
        ttk.Button(frame_file, text="Procurar Lista CSV", command=self.procurar_ficheiro, width=20).pack(side=tk.LEFT, padx=(0, 10))
        ttk.Label(frame_file, textvariable=self.caminho_ficheiro, font=("Segoe UI", 9, "italic"), foreground="gray").pack(side=tk.LEFT, fill=tk.X, expand=True)

        frame_anexos = ttk.Frame(main_frame)
        frame_anexos.grid(row=6, column=0, columnspan=2, sticky="we", pady=5)
        ttk.Button(frame_anexos, text="Opcional: Pasta Anexos", command=self.procurar_pasta_anexos, width=20).pack(side=tk.LEFT, padx=(0, 10))
        ttk.Label(frame_anexos, textvariable=self.caminho_anexos, font=("Segoe UI", 9, "italic"), foreground="gray").pack(side=tk.LEFT, fill=tk.X, expand=True)
        ttk.Label(main_frame, text="⚠️ O nome do ficheiro deve ser igual ao Nome ou E-mail da pessoa (ex: joao@...pdf)", font=("Segoe UI", 8), foreground="#d32f2f").grid(row=7, column=0, columnspan=2, sticky="w")

        # 3. Mensagem
        lbl_mensagem = ttk.Label(main_frame, text="3. A Mensagem", font=("Segoe UI", 11, "bold"))
        lbl_mensagem.grid(row=8, column=0, columnspan=2, sticky="w", pady=(15, 5))

        ttk.Label(main_frame, text="Assunto:").grid(row=9, column=0, sticky="w", pady=2)
        ttk.Entry(main_frame, textvariable=self.assuntoVar, width=60).grid(row=9, column=1, sticky="w", pady=2)

        lbl_dica = ttk.Label(main_frame, text="Dica: Escreva {Nome} no texto para ser substituído automaticamente.", font=("Segoe UI", 9), foreground="blue")
        lbl_dica.grid(row=10, column=0, columnspan=2, sticky="w", pady=2)

        ttk.Label(main_frame, text="Corpo da Mensagem:").grid(row=11, column=0, sticky="nw", pady=5)
        self.txt_corpo = tk.Text(main_frame, height=10, width=50, font=("Segoe UI", 10))
        self.txt_corpo.grid(row=11, column=1, sticky="w", pady=5)

        # Botão Enviar e Progresso
        self.btn_enviar = ttk.Button(main_frame, text="🚀 INICIAR ENVIOS", command=self.iniciar_envio_thread)
        self.btn_enviar.grid(row=12, column=0, columnspan=2, pady=(15, 10))
        
        self.lbl_status = ttk.Label(main_frame, text="A aguardar dados...", font=("Segoe UI", 9), foreground="blue")
        self.lbl_status.grid(row=13, column=0, columnspan=2)

    def procurar_ficheiro(self):
        filename = filedialog.askopenfilename(
            title="Selecione o ficheiro CSV de Contactos",
            filetypes=(("Ficheiros CSV", "*.csv"), ("Ficheiros de Texto", "*.txt"), ("Todos os ficheiros", "*.*"))
        )
        if filename:
            self.caminho_ficheiro.set(filename)
            self.carregar_destinatarios(filename)

    def procurar_pasta_anexos(self):
        dirname = filedialog.askdirectory(title="Selecione a Pasta com os Anexos")
        if dirname:
            self.caminho_anexos.set(dirname)

    def carregar_destinatarios(self, filename):
        self.destinatarios = []
        try:
            with open(filename, "r", encoding="utf-8-sig") as file:
                # Tentar ler como CSV com cabeçalhos
                leitor = csv.DictReader(file)
                if not leitor.fieldnames:
                    raise Exception("Ficheiro vazio ou formato inválido.")
                
                # Normalizar nomes das colunas
                colunas = [c.strip().lower() for c in leitor.fieldnames]
                col_email = None
                col_nome = None
                
                for idx, c in enumerate(colunas):
                    if "mail" in c: col_email = leitor.fieldnames[idx]
                    if "nome" in c or "name" in c: col_nome = leitor.fieldnames[idx]
                
                if not col_email:
                    # Se não encontrou cabeçalho, assumir que é só uma lista de emails ou Nome,Email
                    file.seek(0)
                    for linha in file:
                        partes = linha.strip().split(',')
                        if len(partes) >= 2 and "@" in partes[1]:
                            self.destinatarios.append({'Nome': partes[0].strip(), 'Email': partes[1].strip()})
                        elif "@" in partes[0]:
                            self.destinatarios.append({'Nome': '', 'Email': partes[0].strip()})
                else:
                    for linha in leitor:
                        email = linha.get(col_email, '').strip()
                        nome = linha.get(col_nome, '').strip() if col_nome else ''
                        if email and "@" in email:
                            self.destinatarios.append({'Nome': nome, 'Email': email})
                            
            if self.destinatarios:
                self.lbl_status.config(text=f"✅ {len(self.destinatarios)} destinatário(s) lido(s) com sucesso.", foreground="green")
            else:
                self.lbl_status.config(text="❌ Nenhum e-mail encontrado no ficheiro.", foreground="red")
        except Exception as e:
            messagebox.showerror("Erro de Leitura", f"Não foi possível ler o ficheiro CSV:\n{e}")

    def iniciar_envio_thread(self):
        if not self.emailVar.get() or not self.passVar.get() or not self.assuntoVar.get() or not self.destinatarios:
            messagebox.showwarning("Faltam Dados", "Por favor preencha o E-mail, Password, Assunto e escolha o ficheiro CSV com contactos.")
            return
            
        corpo = self.txt_corpo.get("1.0", tk.END).strip()
        if not corpo:
            messagebox.showwarning("Sem Mensagem", "O corpo do e-mail não pode estar vazio.")
            return

        confirma = messagebox.askyesno("Confirmar Envios", f"Tem a certeza que deseja enviar {len(self.destinatarios)} e-mails individuais?")
        if confirma:
            self.btn_enviar.config(state="disabled", text="A ENVIAR...")
            threading.Thread(target=self.enviar_emails, args=(corpo,), daemon=True).start()

    def enviar_emails(self, corpo_base):
        email_remetente = self.emailVar.get().strip()
        password_app = self.passVar.get().strip()
        assunto_base = self.assuntoVar.get().strip()
        pasta_anexos = self.caminho_anexos.get()
        
        self.lbl_status.config(text="A ligar aos servidores da Google...", foreground="blue")
        
        sucessos = 0
        erros = 0
        
        try:
            server = smtplib.SMTP("smtp.gmail.com", 587)
            server.set_debuglevel(0)
            server.starttls()
            server.login(email_remetente, password_app)
            
            # Listar ficheiros na pasta de anexos (se existir)
            ficheiros_pasta = []
            if pasta_anexos and os.path.isdir(pasta_anexos):
                ficheiros_pasta = os.listdir(pasta_anexos)
            
            for index, pessoa in enumerate(self.destinatarios):
                nome = pessoa['Nome']
                email_dest = pessoa['Email']
                
                self.lbl_status.config(text=f"A enviar {index + 1} de {len(self.destinatarios)}... ({email_dest})")
                self.root.update()
                
                try:
                    msg = MIMEMultipart()
                    msg['From'] = email_remetente
                    msg['To'] = email_dest
                    
                    # Substituir {Nome} no assunto e corpo
                    assunto_final = assunto_base.replace("{Nome}", nome).replace("{nome}", nome)
                    corpo_final = corpo_base.replace("{Nome}", nome).replace("{nome}", nome)
                    
                    msg['Subject'] = assunto_final
                    msg.attach(MIMEText(corpo_final, 'plain', 'utf-8'))
                    
                    # Procurar anexo específico para esta pessoa
                    if ficheiros_pasta:
                        for ficheiro_nome in ficheiros_pasta:
                            nome_sem_extensao = os.path.splitext(ficheiro_nome)[0].strip().lower()
                            # Tenta corresponder ao email ou ao nome completo
                            if nome_sem_extensao == email_dest.lower() or (nome and nome_sem_extensao == nome.lower()):
                                caminho_completo = os.path.join(pasta_anexos, ficheiro_nome)
                                with open(caminho_completo, "rb") as f:
                                    part = MIMEApplication(f.read(), Name=ficheiro_nome)
                                part['Content-Disposition'] = f'attachment; filename="{ficheiro_nome}"'
                                msg.attach(part)
                                break # Anexa só o primeiro que encontrar
                    
                    server.send_message(msg)
                    sucessos += 1
                    time.sleep(1) # Pausa de segurança
                except Exception as e:
                    print(f"Erro ao enviar para {email_dest}: {e}")
                    erros += 1

            server.quit()
            
            self.lbl_status.config(text=f"Concluído: {sucessos} enviados, {erros} erros.", foreground="green")
            messagebox.showinfo("Envios Terminados", f"🔥 Processo concluído!\n\nEnviados com sucesso: {sucessos}\nFalharam: {erros}")
            
        except smtplib.SMTPAuthenticationError:
            self.lbl_status.config(text="Acesso negado: Password ou E-mail inválidos.", foreground="red")
            messagebox.showerror("Erro de Login", "A Google recusou a ligação.\nCertifique-se que o E-mail é correto e usou uma 'App Password' válida de 16 caracteres.")
        except Exception as e:
            self.lbl_status.config(text="Ocorreu um erro central.", foreground="red")
            messagebox.showerror("Erro Crítico", f"Falha na comunicação com a Google:\n{e}")
        finally:
            self.btn_enviar.config(state="normal", text="🚀 INICIAR ENVIOS")

if __name__ == "__main__":
    root = tk.Tk()
    app = EmailSenderApp(root)
    root.mainloop()
