import customtkinter as ctk
from tkinter import *
from tkinter import messagebox
import openpyxl
import pathlib

# Configura a aparência padrão do sistema
ctk.set_appearance_mode("Light")  # Definindo o tema Padrão de inicialização
ctk.set_default_color_theme("green")


class App(ctk.CTk):
    def __init__(self):
        super().__init__()  # Classe principal
        self.layout_config()
        self.appearance()
        self.todo_sistema()

    def layout_config(self):
        self.title("Sistema de Gestão de Clientes")
        self.geometry("700x500")

    def appearance(self):
        # Label para selecionar o tema
        self.lb_apm = ctk.CTkLabel(self, text="Tema", bg_color="transparent")
        self.lb_apm.place(x=50, y=430)

        # OptionMenu para trocar o tema
        self.opt_apm = ctk.CTkOptionMenu(self, values=['Light', 'Dark', 'System'], command=self.change_apm)
        self.opt_apm.place(x=50, y=460)

    def todo_sistema(self):
        # Frame superior
        frame = ctk.CTkFrame(self, width=700, height=50, corner_radius=0, fg_color="teal")
        frame.place(x=0, y=10)
        title = ctk.CTkLabel(frame, text="Sistema de Gestão de Clientes", font=("Arial", 24), text_color="#FFFFFF")
        title.place(x=190, y=10)

        # Instrução para o usuário
        span = ctk.CTkLabel(self, text="Por favor preencha o formulário!", font=("Arial", 16))
        span.place(x=50, y=60)

        # Variáveis de texto
        self.name_value = StringVar()
        self.contact_value = StringVar()
        self.age_value = StringVar()
        self.address_value = StringVar()

        # Entrys
        self.name_entry = ctk.CTkEntry(self, width=350, textvariable=self.name_value, font=("Arial", 16))
        self.contact_entry = ctk.CTkEntry(self, width=200, textvariable=self.contact_value, font=("Arial", 16))
        self.age_entry = ctk.CTkEntry(self, width=150, textvariable=self.age_value, font=("Arial", 16))
        self.address_entry = ctk.CTkEntry(self, width=200, textvariable=self.address_value, font=("Arial", 16))

        # Combobox
        self.gender_combobox = ctk.CTkComboBox(self, values=["Masculino", "Feminino"], font=("Century Gothic bold", 14),
                                               width=150)
        self.gender_combobox.set("Masculino")

        # Entrada de Observações
        self.obs_entry = ctk.CTkTextbox(self, width=465, height=150, font=("Arial", 18), border_width=2)

        # Labels
        lb_name = ctk.CTkLabel(self, text="Nome Completo:", font=("Arial", 16))
        lb_contact = ctk.CTkLabel(self, text="Contato", font=("Arial", 16))
        lb_age = ctk.CTkLabel(self, text="Data de Nascimento", font=("Arial", 16))
        lb_gender = ctk.CTkLabel(self, text="Gênero:", font=("Arial", 16))
        lb_address = ctk.CTkLabel(self, text="Endereço:", font=("Arial", 16))
        lb_obs = ctk.CTkLabel(self, text="Observações:", font=("Arial", 16))

        # Posicionando os elementos na janela
        lb_name.place(x=50, y=120)
        self.name_entry.place(x=50, y=150)
        lb_contact.place(x=450, y=120)
        self.contact_entry.place(x=450, y=150)
        lb_age.place(x=300, y=190)
        self.age_entry.place(x=300, y=220)
        lb_gender.place(x=500, y=190)
        self.gender_combobox.place(x=500, y=220)
        lb_address.place(x=50, y=190)
        self.address_entry.place(x=50, y=220)
        lb_obs.place(x=50, y=260)
        self.obs_entry.place(x=185, y=260)

        # Botões com cores dinâmicas baseadas no tema
        btn_submit = ctk.CTkButton(self, text="SALVAR DADOS", command=self.submit,
                                   fg_color=ctk.ThemeManager.theme["CTkButton"]["fg_color"],
                                   hover_color=ctk.ThemeManager.theme["CTkButton"]["hover_color"])

        btn_submit.place(x=200, y=420)

        btn_clear = ctk.CTkButton(self, text="LIMPAR CAMPOS", command=self.clear,
                              fg_color=ctk.ThemeManager.theme["CTkButton"]["fg_color"],
                              hover_color=ctk.ThemeManager.theme["CTkButton"]["hover_color"])
        btn_clear.place(x=350, y=420)

        btn_delete = ctk.CTkButton(self, text="APAGAR CADASTRO", command=self.delete,
                               fg_color="#C0392B", hover_color="#E74C3C")
        btn_delete.place(x=500, y=420)

    def submit(self):
        ficheiro = pathlib.Path("Clientes.xlsx")

        if not ficheiro.exists():
            # Cria um novo arquivo Excel
            workbook = openpyxl.Workbook()
            folha = workbook.active
            # Define os cabeçalhos na primeira linha
            folha['A1'] = "Nome Completo"
            folha['B1'] = "CONTATO"
            folha['C1'] = "IDADE"
            folha['D1'] = "Gênero"
            folha['E1'] = "Endereço"
            folha['F1'] = "Obs"
            # Salva o arquivo
            workbook.save(ficheiro)

        # Função para salvar os dados
        name = self.name_value.get()
        contact = self.contact_value.get()
        age = self.age_value.get()
        address = self.address_value.get()
        gender = self.gender_combobox.get()
        obs = self.obs_entry.get("1.0", END).strip()

        # Validação simples
        if not name or not contact:
            messagebox.showwarning("Aviso", "Por favor, preencha os campos obrigatórios!")
            return

        # Carrega o arquivo existente
        workbook = openpyxl.load_workbook(ficheiro)
        folha = workbook.active

        # Encontra a próxima linha disponível
        next_row = folha.max_row + 1

        # Preenche os dados na próxima linha disponível
        folha.cell(column=1, row=next_row, value=name)
        folha.cell(column=2, row=next_row, value=contact)
        folha.cell(column=3, row=next_row, value=age)
        folha.cell(column=4, row=next_row, value=gender)
        folha.cell(column=5, row=next_row, value=address)
        folha.cell(column=6, row=next_row, value=obs)

        # Salva as alterações no arquivo
        workbook.save(ficheiro)

        # Exibe uma mensagem de sucesso
        messagebox.showinfo("Sistema", "Dados salvos com sucesso!")
        self.clear()

    def clear(self):
        # Função para limpar os campos
        self.name_value.set("")
        self.contact_value.set("")
        self.age_value.set("")
        self.address_value.set("")
        self.gender_combobox.set("Masculino")
        self.obs_entry.delete('0.0', END)

    def delete(self):
        ficheiro = pathlib.Path("Clientes.xlsx")
        if ficheiro.exists():
            # Carrega o arquivo existente
            workbook = openpyxl.load_workbook(ficheiro)
            folha = workbook.active

            # Obtém o nome para exclusão
            name = self.name_value.get()

            for row in range(2, folha.max_row + 1):
                if folha.cell(row=row, column=1).value == name:
                    folha.delete_rows(row)
                    workbook.save(ficheiro)
                    messagebox.showinfo("Sistema", f"Registro de {name} excluído com sucesso!")
                    self.clear()
                    return

            messagebox.showinfo("Sistema", "Registro não encontrado.")
        else:
            messagebox.showwarning("Sistema", "Arquivo de clientes não encontrado.")

    def change_apm(self, new_appearance):
        ctk.set_appearance_mode(new_appearance)
        self.update_colors()

    def update_colors(self):
        # Atualiza as cores dos elementos conforme o tema
        current_theme = ctk.get_appearance_mode()
        if current_theme == "Dark":
            text_color = "#FFFFFF"
        else:
            text_color = "#000000"

        # Atualiza as labels
        for widget in self.winfo_children():
            if isinstance(widget, ctk.CTkLabel):
                widget.configure(text_color=text_color)

        # Atualiza os botões, se necessário
        # Nota: Se os botões usarem cores dinâmicas, talvez não seja necessário atualizar

        # Atualiza outros elementos que precisam de cores específicas
        # Por exemplo, o título no frame superior já está com texto branco, que funciona bem em ambos os temas


if __name__ == "__main__":
    app = App()
    app.mainloop()
