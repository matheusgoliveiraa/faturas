import pandas as pd
import tkinter as tk
from tkinter import ttk, messagebox
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import smtplib
from PIL import Image, ImageTk

# dentro dos parenteses insira a sua planilha excel. 
def load_data_from_excel():
    df = pd.read_excel()
    df['Unidades'] = df['Unidades'].fillna('Não tem unidade')
    return df

# Função para carregar usuários únicos
def load_usuarios(df):
    return df['Usuario'].unique().tolist()

# Função para carregar empresas únicas com base no usuário selecionado
def load_empresas(df, usuario):
    if usuario:
        empresas = df[df['Usuario'] == usuario]
        return empresas['Empresas'].unique().tolist()
    return []

# Função para carregar unidades com base na empresa selecionada
def load_unidades(df, usuario, empresa):
    unidades = df[(df['Usuario'] == usuario) & (df['Empresas'] == empresa)]
    return unidades[['Unidades', 'Emails']].drop_duplicates().values.tolist()

def send_email(cliente, unidade, meses_selecionados, ano, emails):
    smtp_server = '' # insira a configuração do provedor de e-mail.
    smtp_port = # porta
    smtp_user = '' # insira seu e-mail
    smtp_password = '' # insira sua senha

    # Configuração do e-mail
    from_email = '' # repita o mesmo e-mail do smtp_user
    to_emails = emails
    subject = '' # titulo do e-mail

    # Construção do corpo do e-mail
    body = f"""
    Prezados,

    Espero que estejam bem.

    Não identificamos o recebimento de algumas faturas de unidades,
    Peço a gentileza de nos encaminhar conforme abaixo:

    Cliente: {cliente}
    """
    if unidade and unidade != 'Não tem unidade':
        body += f"Unidade: {unidade}\n"
    body += f"    Meses: {', '.join(meses_selecionados)}\n"
    body += f"    Ano: {ano}\n"

    body += """
    Por favor, efetue o envio o mais breve possível.

    Atenciosamente,
    """

    # Criação do e-mail
    msg = MIMEMultipart()
    msg['From'] = from_email
    msg['To'] = ', '.join(to_emails)
    msg['Subject'] = subject
    msg.attach(MIMEText(body.strip(), 'plain'))  # Remover espaços extras

    # Envio do e-mail
    try:
        with smtplib.SMTP(smtp_server, smtp_port) as server:
            server.starttls()
            server.login(smtp_user, smtp_password)
            server.sendmail(from_email, to_emails, msg.as_string())
        print("E-mail enviado com sucesso.")
    except Exception as e:
        print(f"Erro ao enviar e-mail: {e}")

def submit_action():
    usuario = entry_usuario.get()
    cliente = entry_cliente.get()
    unidade = entry_unidade.get()
    mes_indices = listbox_mes.curselection()
    meses_selecionados = [listbox_mes.get(i) for i in mes_indices]
    ano = entry_ano.get()

    # Validação dos campos
    if not usuario:
        messagebox.showerror("Erro", "Por favor, selecione um front/middle.")
        return
    if not cliente:
        messagebox.showerror("Erro", "Por favor, selecione um cliente.")
        return
    if unidade != 'Não tem unidade' and not unidade:
        messagebox.showerror("Erro", "Por favor, selecione uma unidade.")
        return
    if not meses_selecionados:
        messagebox.showerror("Erro", "Por favor, selecione pelo menos um mês.")
        return
    if not ano:
        messagebox.showerror("Erro", "Por favor, selecione um ano.")
        return

    print(f"Front/Middle: {usuario}")
    print(f"Cliente: {cliente}")
    print(f"Unidade: {unidade if unidade else 'Não especificada'}")
    print(f"Meses: {meses_selecionados}")
    print(f"Ano: {ano}")

    # Obter e-mails da unidade ou cliente
    if unidade and unidade != 'Não tem unidade':
        unidade_data = data_frame[(data_frame['Usuario'] == usuario) & (data_frame['Empresas'] == cliente) & (data_frame['Unidades'] == unidade)]
    else:
        unidade_data = data_frame[(data_frame['Usuario'] == usuario) & (data_frame['Empresas'] == cliente) & (data_frame['Unidades'].isna() | (data_frame['Unidades'] == 'Não tem unidade'))]

    if unidade_data.shape[0] > 0:
        emails = unidade_data['Emails'].values[0].split(',')
    else:
        emails = []

    # Verificação de e-mails disponíveis
    if not emails or 'Não tem e-mail' in emails:
        messagebox.showerror("Erro", "Não há e-mails disponíveis para o cliente selecionado.")
        return

    # Enviar e-mail com os dados
    send_email(cliente, unidade if unidade and unidade != 'Não tem unidade' else None, meses_selecionados, ano, emails)


def update_empresas(event):
    usuario = entry_usuario.get()
    if usuario:
        empresas = load_empresas(data_frame, usuario)
        entry_cliente['values'] = empresas
    else:
        entry_cliente['values'] = []
    entry_cliente.set('')
    entry_unidade['values'] = []
    entry_emails['values'] = []

def update_unidades(event):
    usuario = entry_usuario.get()
    empresa = entry_cliente.get()
    if usuario and empresa:
        unidades = load_unidades(data_frame, usuario, empresa)
        unidades_names = [u[0] for u in unidades if u[0] != '']  # Exclui entradas vazias
        if not unidades_names:
            unidades_names.append('Não tem unidade')  # Adiciona a opção se não houver unidades

        # Remove duplicatas e garante que 'Não tem unidade' apareça apenas uma vez
        unidades_names = list(dict.fromkeys(unidades_names))
        
        entry_unidade['values'] = unidades_names
    else:
        entry_unidade['values'] = []
    entry_unidade.set('')  # Limpa o valor atual do Combobox de unidades
    entry_emails['values'] = []  # Limpa o Combobox de e-mails

def update_emails(event):
    usuario = entry_usuario.get()
    empresa = entry_cliente.get()
    unidade = entry_unidade.get()
    if usuario and empresa:
        unidades_df = pd.DataFrame(load_unidades(data_frame, usuario, empresa), columns=['Unidades', 'Emails'])
        email_list = unidades_df[unidades_df['Unidades'] == unidade]['Emails'].values

        if email_list.size > 0:
            if pd.isna(email_list[0]) or email_list[0].strip() == '':
                emails = ['Não tem e-mail']
            else:
                emails = email_list[0].split(',')
        else:
            emails = ['Não tem e-mail']
        
        emails = list(dict.fromkeys(emails))
        entry_emails['values'] = emails
        entry_emails.set('')

        if unidade == 'Não tem unidade' and emails[0] == 'Não tem e-mail':
            # Se a unidade for 'Não tem unidade' e não houver e-mails, desativa os campos subsequentes
            entry_emails.config(state=tk.DISABLED)
            entry_ano.config(state=tk.DISABLED)
            listbox_mes.config(state=tk.DISABLED)
        else:
            # Ativa os campos subsequentes
            entry_emails.config(state=tk.NORMAL)
            entry_ano.config(state=tk.NORMAL)
            listbox_mes.config(state=tk.NORMAL)

def update_mes_selecionados(event=None):
    mes_indices = listbox_mes.curselection()
    meses_selecionados = [listbox_mes.get(i) for i in mes_indices]
    meses_texto = ', '.join(meses_selecionados)
    
    if meses_selecionados:
        entry_mes_selecionados.config(state=tk.NORMAL)
        entry_mes_selecionados.delete(0, tk.END)
        entry_mes_selecionados.insert(0, meses_texto)
        entry_mes_selecionados.config(state=tk.DISABLED)
    else:
        entry_mes_selecionados.config(state=tk.NORMAL)
        entry_mes_selecionados.delete(0, tk.END)
        entry_mes_selecionados.config(state=tk.DISABLED)
    
    # Ajusta a largura do campo de entrada com base no comprimento do texto
    texto_largura = len(meses_texto)  # Comprimento do texto dos meses selecionados
    entry_mes_selecionados.config(width=max(20, texto_largura))

def clear_selection():
    # Limpa a seleção do Combobox de usuário
    entry_usuario.set('')
    
    # Limpa a seleção do Combobox de cliente
    entry_cliente.set('')
    
    # Limpa a seleção do Combobox de unidade
    entry_unidade.set('')
    
    # Limpa a seleção do Combobox de e-mails
    entry_emails.set('')
    
    # Limpa a seleção da Listbox de meses
    listbox_mes.selection_clear(0, tk.END)
    
    # Limpa o campo de meses selecionados
    entry_mes_selecionados.config(state=tk.NORMAL)
    entry_mes_selecionados.delete(0, tk.END)
    entry_mes_selecionados.config(state=tk.DISABLED)
    
    # Limpa o Combobox de ano
    entry_ano.set('')

# Caminho para o arquivo Excel
file_path = ''

# Ler os dados da planilha
data_frame = load_data_from_excel(file_path)

# Definindo cores e fontes
bg_color = '#292F36'
font_family = 'Roboto, sans-serif'
font_size = 11
font_color = '#FF8C42'

# Criando a janela principal
root = tk.Tk()
root.title("Formulário")
root.configure(bg=bg_color)

# Definindo as colunas do grid para redimensionamento
root.columnconfigure(0, weight=1)
root.columnconfigure(1, weight=3)
for i in range(9):
    root.rowconfigure(i, weight=1)

# Carregando o logotipo
logo_path = ''  # Substitua pelo caminho do seu logotipo
logo_image = Image.open(logo_path)
logo_image = logo_image.resize((100, 100), Image.LANCZOS)
logo_photo = ImageTk.PhotoImage(logo_image)

# Inserindo o logotipo
logo_label = tk.Label(root, image=logo_photo, bg=bg_color)
logo_label.grid(column=0, row=0, padx=10, pady=10, columnspan=2)

# Criação dos rótulos e campos de entrada com cores e fontes
style = ttk.Style()
style.configure('TLabel', background=bg_color, font=(font_family, font_size), foreground=font_color)
style.configure('TCombobox', font=(font_family, font_size), foreground=font_color)
style.configure('TButton', font=(font_family, font_size), foreground=font_color)
# Configuração do estilo para os Comboboxes e Botões
style.configure('TCombobox', font=(font_family, font_size), foreground='black', background='white')
style.configure('TButton', font=(font_family, font_size), foreground='black')  # Texto dos botões em preto

ttk.Label(root, text="Front/Middle:").grid(column=0, row=1, padx=10, pady=10, sticky='ew')
entry_usuario = ttk.Combobox(root, values=load_usuarios(data_frame))
entry_usuario.grid(column=1, row=1, padx=10, pady=10, sticky='ew')
entry_usuario.bind("<<ComboboxSelected>>", update_empresas)

ttk.Label(root, text="Cliente:").grid(column=0, row=2, padx=10, pady=10, sticky='ew')
entry_cliente = ttk.Combobox(root)
entry_cliente.grid(column=1, row=2, padx=10, pady=10, sticky='ew')
entry_cliente.bind("<<ComboboxSelected>>", update_unidades)

ttk.Label(root, text="Unidade:").grid(column=0, row=3, padx=10, pady=10, sticky='ew')
entry_unidade = ttk.Combobox(root)
entry_unidade.grid(column=1, row=3, padx=10, pady=10, sticky='ew')
entry_unidade.bind("<<ComboboxSelected>>", update_emails)

ttk.Label(root, text="E-mails:").grid(column=0, row=4, padx=10, pady=10, sticky='ew')
entry_emails = ttk.Combobox(root)
entry_emails.grid(column=1, row=4, padx=10, pady=10, sticky='ew')

ttk.Label(root, text="Ano:").grid(column=0, row=5, padx=10, pady=10, sticky='ew')
anos = [str(ano) for ano in range(2024, 2051)]
entry_ano = ttk.Combobox(root, values=anos)
entry_ano.grid(column=1, row=5, padx=10, pady=10, sticky='ew')

ttk.Label(root, text="Meses:").grid(column=0, row=6, padx=10, pady=10, sticky='ew')
listbox_mes = tk.Listbox(root, selectmode=tk.MULTIPLE, font=(font_family, font_size))
listbox_mes.grid(column=1, row=6, padx=10, pady=10, sticky='ew')
meses = ["Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho", "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"]
for mes in meses:
    listbox_mes.insert(tk.END, mes)
listbox_mes.bind('<<ListboxSelect>>', update_mes_selecionados)

entry_mes_selecionados = ttk.Entry(root, state=tk.DISABLED, font=(font_family, font_size))
entry_mes_selecionados.grid(column=1, row=7, padx=10, pady=10, sticky='ew')

# Botão de Submissão
submit_button = ttk.Button(root, text="Enviar", command=submit_action)
submit_button.grid(column=1, row=8, padx=10, pady=10, sticky='ew')

# Botão de Limpar Seleção
clear_button = ttk.Button(root, text="Limpar Seleção", command=clear_selection)
clear_button.grid(column=0, row=8, padx=10, pady=10, sticky='ew')

root.mainloop()
