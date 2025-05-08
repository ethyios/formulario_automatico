import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import os
import sys # Adicionado
import importlib # Adicionado para recarregar no gerador principal

# Nome da pasta para armazenar o arquivo de bytes do modelo
HIDDEN_FOLDER_NAME = "._modelo_data"
# Nome base do arquivo Python que armazenará os bytes do DOCX
OUTPUT_PYTHON_FILENAME_BASENAME = "declaracao_base_bytes.py"

PLACEHOLDERS_TEXT = """O arquivo .docx selecionado deve ser um modelo e precisa conter os seguintes placeholders (marcadores) que serão substituídos:
- {{NOME_RESPONSAVEL}}: Nome completo do(a) responsável.
- {{NOME_FILHO}}: Nome completo do filho(a).
- {{SERIE}}: Série/ano escolar do filho(a).
- {{DATA}}: Data da declaração (será formatada por extenso).
- {{PERIODO}}: Período em que o aluno esteve presente (ex: manhã, tarde, integral).

Exemplo: Prezada {{NOME_RESPONSAVEL}}, declaramos que {{NOME_FILHO}} esteve presente no período da {{PERIODO}}...

Certifique-se de que o arquivo DOCX não esteja protegido ou corrompido.
"""

def get_application_path():
    """Retorna o caminho base da aplicação, seja script ou executável."""
    if getattr(sys, 'frozen', False): # Rodando como bundle PyInstaller
        # sys.executable é o caminho para o .exe
        return os.path.dirname(sys.executable)
    else: # Rodando como script
        # __file__ é o caminho para este script (importar_declaracao.py)
        # Se este script for importado, este será o seu diretório.
        return os.path.dirname(os.path.abspath(__file__))


def gerar_arquivo_python_com_bytes(docx_filepath, original_filename):
    """
    Lê o arquivo .docx e cria um arquivo Python (.py)
    contendo os bytes do documento em uma variável, no subdiretório HIDDEN_FOLDER_NAME da aplicação.
    Retorna True se sucesso, False caso contrário.
    """
    try:
        with open(docx_filepath, "rb") as f:
            content_bytes = f.read()

        conteudo_py = f"# Arquivo gerado automaticamente, não edite manualmente.\n"
        conteudo_py += f"# Contém os bytes do arquivo original: {os.path.basename(original_filename)}\n"
        conteudo_py += f"DOCX_BYTES = {repr(content_bytes)}\n"

        app_path = get_application_path()
        hidden_dir_path = os.path.join(app_path, HIDDEN_FOLDER_NAME)
        
        # Cria o diretório oculto se não existir
        os.makedirs(hidden_dir_path, exist_ok=True)
        
        output_py_path = os.path.join(hidden_dir_path, OUTPUT_PYTHON_FILENAME_BASENAME)

        with open(output_py_path, "w", encoding="utf-8") as f_py:
            f_py.write(conteudo_py)
        return True
    except FileNotFoundError:
        messagebox.showerror("Erro", f"O arquivo '{os.path.basename(original_filename)}' não foi encontrado no caminho especificado: {docx_filepath}", icon='error')
        return False
    except PermissionError:
        app_path = get_application_path() # Recalcula para a mensagem de erro
        hidden_dir_path = os.path.join(app_path, HIDDEN_FOLDER_NAME) # Adicionado para clareza na mensagem
        output_py_path = os.path.join(hidden_dir_path, OUTPUT_PYTHON_FILENAME_BASENAME)
        messagebox.showerror("Erro de Permissão",
                             f"Não foi possível criar o arquivo '{OUTPUT_PYTHON_FILENAME_BASENAME}' em '{hidden_dir_path}'.\n"
                             "Verifique as permissões de escrita no diretório.", icon='error')
        return False
    except Exception as e:
        messagebox.showerror("Erro Inesperado", f"Ocorreu um erro inesperado ao processar o arquivo: {e}", icon='error')
        return False

class ImportadorApp:
    def __init__(self, master):
        self.master = master
        master.title("Importador de Modelo DOCX")
        # master.geometry("550x350") # Ajustar tamanho conforme necessário

        self.frame = ttk.Frame(master, padding="20")
        self.frame.pack(expand=True, fill=tk.BOTH)

        ttk.Label(self.frame, text="Instruções para o Modelo DOCX:", font=("Arial", 12, "bold")).pack(pady=(0,5), anchor=tk.W)

        self.text_instrucoes = tk.Text(self.frame, wrap=tk.WORD, height=12, width=65, relief=tk.SOLID, borderwidth=1, font=("Arial", 9))
        self.text_instrucoes.insert(tk.END, PLACEHOLDERS_TEXT)
        self.text_instrucoes.config(state=tk.DISABLED) # Tornar não editável
        self.text_instrucoes.pack(pady=(0,15), fill=tk.X, expand=True)


        self.btn_selecionar = ttk.Button(self.frame, text="Selecionar Arquivo .docx e Gerar Bytes", command=self.selecionar_e_processar_arquivo)
        self.btn_selecionar.pack(pady=10, fill=tk.X, ipady=5)

        self.status_label = ttk.Label(self.frame, text="Aguardando seleção do arquivo modelo...")
        self.status_label.pack(pady=(5,0), anchor=tk.W)
        
        master.update_idletasks() # Garante que a janela se ajuste ao conteúdo
        master.minsize(master.winfo_width(), master.winfo_height())


    def selecionar_e_processar_arquivo(self):
        self.status_label.config(text="Aguardando seleção...")
        filepath = filedialog.askopenfilename(
            title="Selecione o arquivo DOCX modelo",
            filetypes=(("Word Documents", "*.docx"), ("All files", "*.*"))
        )
        if filepath:
            original_filename = os.path.basename(filepath)
            self.status_label.config(text=f"Processando '{original_filename}'...")
            self.master.update_idletasks()

            if gerar_arquivo_python_com_bytes(filepath, original_filename):
                messagebox.showinfo("Sucesso",
                                    f"Arquivo '{OUTPUT_PYTHON_FILENAME_BASENAME}' gerado/atualizado com sucesso em '{HIDDEN_FOLDER_NAME}' com os bytes de '{original_filename}'.\n\n"
                                    f"Você pode fechar esta janela.\n"
                                    f"O gerador principal tentará recarregar o modelo automaticamente.", icon='info')
                self.status_label.config(text=f"'{OUTPUT_PYTHON_FILENAME_BASENAME}' gerado em '{HIDDEN_FOLDER_NAME}'!")
                # self.master.destroy() # Fecha a janela do importador automaticamente
            else:
                # Mensagem de erro já é mostrada por gerar_arquivo_python_com_bytes
                self.status_label.config(text="Falha ao gerar arquivo de bytes. Verifique a mensagem de erro.")
        else:
            self.status_label.config(text="Nenhum arquivo selecionado. Operação cancelada.")

def iniciar_interface_importador():
    """Função para ser chamada pelo script principal para iniciar esta GUI."""
    root_importador = tk.Toplevel() # Usar Toplevel se chamado de outra GUI Tkinter
    root_importador.grab_set() # Torna a janela modal em relação à janela pai
    
    # Tenta usar um tema mais moderno se disponível
    style = ttk.Style(root_importador) # Aplica estilo ao root_importador
    # Verifica temas disponíveis e tenta usar um mais moderno
    # Temas comuns: 'clam', 'alt', 'default', 'classic'
    # Em Windows, 'vista' pode estar disponível e ser mais moderno que 'clam' às vezes.
    available_themes = style.theme_names()
    if 'vista' in available_themes:
        style.theme_use('vista')
    elif 'clam' in available_themes:
        style.theme_use('clam')

    app = ImportadorApp(root_importador)
    root_importador.wait_window() # Espera esta janela ser fechada antes de retornar

# Este script agora é um módulo, então o if __name__ == "__main__": é removido.
# Se precisar testar este módulo isoladamente, pode adicionar temporariamente:
# if __name__ == "__main__":
#     # Para teste, criar um root principal simples se não houver um
#     root = tk.Tk()
#     root.withdraw() # Esconde a janela root principal de teste
#     iniciar_interface_importador()
#     root.mainloop() # Necessário se root_importador não for Toplevel ou não usar wait_window