import sys
import sqlite3
import tkinter as tk
import matplotlib.pyplot as plt
import numpy as np
import datetime
from tkinter import messagebox, simpledialog

print(repr(r'Valide no debug se o caminho está correto'))

# Conecta ao banco de dados ou cria se não existir
conn = sqlite3.connect('BDExclude.db')

# Cria um cursor para executar comandos SQL
cursor = conn.cursor()

# Cria a tabela BDExclude
cursor.execute("""
CREATE TABLE IF NOT EXISTS BDExclude (
    Justificativa TEXT,
    serial TEXT
);
""")

# Commita as mudanças e fecha a conexão
conn.commit()
conn.close()

class DatabaseManager:
    def __init__(self):
        self.conn = sqlite3.connect(r'Insira o caminho do banco de dados aqui')
        self.cursor = self.conn.cursor()

    def delete_data(self, serial):
        self.cursor.execute("DELETE FROM lavagem_placas WHERE serial = ?", (serial,))
        self.conn.commit()

    def fetch_data(self, serial):
        self.cursor.execute("SELECT * FROM lavagem_placas WHERE serial = ?", (serial,))
        return self.cursor.fetchone()

    def insert_modelo(self, modelo):
        existing_modelos = self.fetch_all_modelos()
        if modelo not in existing_modelos:
            self.cursor.execute("INSERT INTO ID_MODELO (modelo) VALUES (?)", (modelo,))
            self.conn.commit()
            return True
        else:
            return False
        self.conn.commit()

    def fetch_all_modelos(self):
        self.cursor.execute("SELECT modelo FROM ID_MODELO")
        return [row[0] for row in self.cursor.fetchall()] 

    def update_modelo(self, old_modelo, new_modelo):
        self.cursor.execute("UPDATE ID_MODELO SET modelo = ? WHERE modelo = ?", (new_modelo, old_modelo))
        self.conn.commit()

    def delete_modelo(self, modelo):
        self.cursor.execute("DELETE FROM ID_MODELO WHERE modelo = ?", (modelo,))
        self.conn.commit()  

    def insert_user(self, nome):
        self.cursor.execute("INSERT INTO ID_USER (nome) VALUES (?)", (nome,))
        self.conn.commit()

    def update_user(self, old_nome, new_nome):
        self.cursor.execute("UPDATE ID_USER SET nome = ? WHERE nome = ?", (new_nome, old_nome))
        self.conn.commit()

    def delete_user(self, nome):
        self.cursor.execute("DELETE FROM ID_USER WHERE nome = ?", (nome,))
        self.conn.commit()
    

    def fetch_all_users(self):
        self.cursor.execute("SELECT nome FROM ID_USER")
        return [row[0] for row in self.cursor.fetchall()]
    

    def close(self):
        self.conn.close()

class DeleteRecordWindow:
    def __init__(self, root):
        self.root = root
        self.root.title("CLP - ADMIN")
        self.root.geometry("800x500")
        self.root.configure(bg="#4F4F4F")  # Cor de fundo


        self.title_label = tk.Label(root, text="Control Board Admin", font=("Arial", 20, "bold"))
        self.title_label.grid(row=0, column=0, columnspan=4, pady=20, sticky='ew')

        # Agora com espaçamento e alinhamento
        grid_opts = {"padx": 20, "pady": 20, "sticky": "nsew"}

         # Estilos de botão personalizados
        blue_button_style = {"padx": 5, "pady": 10, "bg": "#06098C", "fg": "white", "font": ("Arial", 10)}
        red_button_style = {"padx": 5, "pady": 10, "bg": "red", "fg": "white", "font": ("Arial", 10)}
        yellow_button_style = {"padx": 5, "pady": 10, "bg": "yellow", "fg": "black", "font": ("Arial", 10)}
        green_button_style = {"padx": 5, "pady": 10, "bg": "green", "fg": "white", "font": ("Arial", 10)}

        # Agora com espaçamento e alinhamento
        grid_opts = {"padx": 20, "pady": 20, "sticky": "nsew"}

        self.excluir_usuario_button = tk.Button(root, text="EXCLUIR USUÁRIO", command=self.excluir_usuario, **red_button_style)
        self.excluir_usuario_button.grid(row=1, column=0, **grid_opts)

        self.cadastrar_usuario_button = tk.Button(root, text="CADASTRAR USUÁRIO", command=self.cadastrar_usuario, **green_button_style)
        self.cadastrar_usuario_button.grid(row=1, column=1, **grid_opts)

        self.ver_usuarios_button = tk.Button(root, text="VER LISTA DE USUÁRIOS", command=self.ver_usuarios, **yellow_button_style)
        self.ver_usuarios_button.grid(row=1, column=2, **grid_opts)

        self.editar_usuario_button = tk.Button(root, text="EDITAR USUÁRIO", command=self.editar_usuario, **blue_button_style)
        self.editar_usuario_button.grid(row=1, column=3, **grid_opts)

        self.excluir_modelo_button = tk.Button(root, text="EXCLUIR MODELO", command=self.excluir_modelo, **red_button_style)
        self.excluir_modelo_button.grid(row=2, column=0, **grid_opts)

        self.cadastrar_modelo_button = tk.Button(root, text="CADASTRAR MODELO", command=self.cadastrar_modelo, **green_button_style)
        self.cadastrar_modelo_button.grid(row=2, column=1, **grid_opts)

        self.ver_modelos_button = tk.Button(root, text="VER MODELOS", command=self.ver_modelos, **yellow_button_style)
        self.ver_modelos_button.grid(row=2, column=2, **grid_opts)

        self.alterar_modelo_button = tk.Button(root, text="ALTERAR MODELO", command=self.alterar_modelo, **blue_button_style)
        self.alterar_modelo_button.grid(row=2, column=3, **grid_opts)

        self.delete_button = tk.Button(root, text="EXCLUIR REGISTRO", command=self.delete_record, **red_button_style)
        self.delete_button.grid(row=3, column=0, **grid_opts)

        self.db_manager = DatabaseManager()

    def cadastrar_modelo(self):
        modelo = simpledialog.askstring("Cadastrar Modelo", "Insira o nome do modelo:")
        if modelo:
            modelo = modelo.upper()
            success = self.db_manager.insert_modelo(modelo)
            if success:
                messagebox.showinfo("Modelo Cadastrado", "Modelo cadastrado com sucesso!")
            else:
                messagebox.showwarning("Modelo Duplicado", "Esse modelo já foi cadastrado.")

    def ver_modelos(self):
        filtro = simpledialog.askstring("Filtrar Modelos", "Insira o nome do modelo para filtrar (deixe em branco para ver todos):")
        if filtro:
            filtro = filtro.lower()
        modelos = self.db_manager.fetch_all_modelos()
        if filtro:
            modelos = [modelo for modelo in modelos if filtro in modelo.lower()]
        if modelos:
            modelos_str = "\n".join(modelos)
            messagebox.showinfo("Modelos Cadastrados", f"Modelos:\n{modelos_str}")
        else:
            messagebox.showwarning("Nenhum Modelo Encontrado", "Nenhum modelo encontrado com esse filtro.") 

    def alterar_modelo(self):
        modelos = self.db_manager.fetch_all_modelos()
        if not modelos:
            messagebox.showwarning("Nada pra Alterar", "Não tem nenhum modelo cadastrado pra alterar.")
            return
        old_modelo = simpledialog.askstring("Alterar Modelo", "Insira o nome do modelo que deseja alterar:")
        if old_modelo:
            if old_modelo.upper() not in modelos:
                messagebox.showwarning("Modelo Inexistente", "Esse modelo não foi cadastrado.")
                return
            new_modelo = simpledialog.askstring("Novo Modelo", "Insira o novo nome para o modelo:")
            if new_modelo:
                self.db_manager.update_modelo(old_modelo.upper(), new_modelo.upper())
                messagebox.showinfo("Modelo Alterado", "Modelo alterado com sucesso!")


    def excluir_modelo(self):
        modelos = self.db_manager.fetch_all_modelos()
        if not modelos:
            messagebox.showwarning("Nada pra Excluir", "Não tem nenhum modelo cadastrado pra excluir.")
            return
        modelo = simpledialog.askstring("Excluir Modelo", "Insira o nome do modelo que deseja excluir:")
        if modelo:
            confirmation = messagebox.askyesno("Excluir Modelo", f"Deseja excluir o modelo {modelo.upper()}?")
            if confirmation:
                self.db_manager.delete_modelo(modelo.upper())
                messagebox.showinfo("Modelo Excluído", "Modelo excluído com sucesso!") 

    def cadastrar_usuario(self):
        nome = simpledialog.askstring("Cadastrar Usuário", "Insira o nome do usuário:")
        if nome:
            nome = nome.upper()
            usuarios_existentes = self.db_manager.fetch_all_users()
            if nome not in [usuario[1].upper() for usuario in usuarios_existentes]:
                self.db_manager.insert_user(nome)
                messagebox.showinfo("Usuário Cadastrado", "Usuário cadastrado com sucesso!")
            else:
                messagebox.showwarning("Usuário Duplicado", "Esse usuário já foi cadastrado.")

    def ver_usuarios(self):
        usuarios = self.db_manager.fetch_all_users()
        if usuarios:
            usuarios_str = "\n".join(usuarios)
            messagebox.showinfo("Lista de Usuários", f"Usuários:\n{usuarios_str}")
        else:
            messagebox.showwarning("Nenhum Usuário Encontrado", "Nenhum usuário cadastrado.") 

    def editar_usuario(self):
        usuarios = self.db_manager.fetch_all_users()
        if not usuarios:
            messagebox.showwarning("Nada pra Editar", "Não tem nenhum usuário cadastrado pra editar.")
            return
        old_nome = simpledialog.askstring("Editar Usuário", "Insira o nome do usuário que deseja editar:")
        if old_nome:
            if old_nome.upper() not in usuarios:
                messagebox.showwarning("Usuário Inexistente", "Esse usuário não existe.")
                return
            new_nome = simpledialog.askstring("Novo Nome", "Insira o novo nome para o usuário:")
            if new_nome:
                self.db_manager.update_user(old_nome.upper(), new_nome.upper())
                messagebox.showinfo("Usuário Editado", "Usuário editado com sucesso!")


    def excluir_usuario(self):
        usuarios = self.db_manager.fetch_all_users()
        if not usuarios:
            messagebox.showwarning("Nada pra Excluir", "Não tem nenhum usuário cadastrado pra excluir.")
            return
        nome = simpledialog.askstring("Excluir Usuário", "Insira o nome do usuário que deseja excluir:")
        if nome:
            confirmation = messagebox.askyesno("Excluir Usuário", f"Deseja excluir o usuário {nome.upper()}?")
            if confirmation:
                self.db_manager.delete_user(nome.upper())
                messagebox.showinfo("Usuário Excluído", "Usuário excluído com sucesso!")                                          

    def delete_record(self):
        serial = simpledialog.askstring("Excluir Registro", "Insira o serial:")
        if serial == "SN":
            self.show_sn_records()
        elif serial:
            self.db_manager.cursor.execute("SELECT * FROM lavagem_placas WHERE serial = ?", (serial,))
            records = self.db_manager.cursor.fetchall()         
            if len(records) > 1:
                self.show_duplicate_records(serial, records)
            elif len(records) == 1:
                record = records[0]
                justification = simpledialog.askstring("Justificativa", "Insira a justificativa para exclusão:")
                if justification:
                    details = f"Serial: {record[0]}\nData: {record[1]}\nTurno: {record[2]}\nHora: {record[3]}"
                    confirmation = messagebox.askyesno("Excluir Registro", f"Deseja excluir o seguinte registro?\n\n{details}")
                    if confirmation:
                        conn = sqlite3.connect('BDExclude.db')
                        cursor = conn.cursor()
                        cursor.execute("INSERT INTO BDExclude (Justificativa, serial) VALUES (?, ?)", (justification, serial))
                        conn.commit()
                        conn.close()
                        self.db_manager.delete_data(serial)
                        messagebox.showinfo("Registro Excluído", "Registro excluído com sucesso!")
                else:
                    messagebox.showwarning("Justificativa Necessária", "Por favor, insira uma justificativa para a exclusão.")
            else:
                messagebox.showwarning("Registro Não Encontrado", "Nenhum registro encontrado com esse serial.")

    def show_duplicate_records(self, serial, records):
        dup_window = tk.Toplevel(self.root)
        dup_window.title(f"Registros Duplicados - {serial}")
        dup_window.geometry("950x400")
        dup_window.configure(bg="#4F4F4F")
        
        frame = tk.Frame(dup_window)
        frame.pack(fill=tk.BOTH, expand=1)
        
        scrollbar = tk.Scrollbar(frame, orient=tk.VERTICAL)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        dup_listbox = tk.Listbox(frame, selectmode=tk.MULTIPLE, yscrollcommand=scrollbar.set, font=("Arial", 12), bg="#333", fg="white")
        dup_listbox.pack(fill=tk.BOTH, expand=1)
        
        scrollbar.config(command=dup_listbox.yview)
        
        for record in records:
            dup_listbox.insert(tk.END, f"ID: {record[7]}, Data: {record[1]}, Turno: {record[2]}, Hora: {record[3]}, Responsavel: {record[5]}, Linha: {record[6]}")
        
        def delete_selected_dup():
            selected_indices = dup_listbox.curselection()
            selected_records = [records[i] for i in selected_indices]
            
            justification = simpledialog.askstring("Justificativa", "Insira a justificativa para exclusão:")
            if justification:
                for record in selected_records:
                    self.db_manager.cursor.execute("DELETE FROM lavagem_placas WHERE id = ?", (record[0],))
                    self.db_manager.conn.commit()
                    
                    # Inserir justificativa no BDExclude
                    conn = sqlite3.connect('BDExclude.db')
                    cursor = conn.cursor()
                    cursor.execute("INSERT INTO BDExclude (Justificativa, serial) VALUES (?, ?)", (justification, serial))
                    conn.commit()
                    conn.close()
                    
                dup_window.destroy()
                messagebox.showinfo("Registros Excluídos", f"{len(selected_records)} registros com serial {serial} foram excluídos.")
            else:
                messagebox.showwarning("Justificativa Necessária", "Por favor, insira uma justificativa para a exclusão.")
        
        delete_button = tk.Button(dup_window, text="Excluir Selecionados", command=delete_selected_dup)
        delete_button.pack()

    def show_sn_records(self):
        sn_window = tk.Toplevel(self.root)
        sn_window.title("Registros SN")
        sn_window.geometry("950x400")
        sn_window.configure(bg="#4F4F4F")  # Cor de fundo igual à janela principal
        
        frame = tk.Frame(sn_window)
        frame.pack(fill=tk.BOTH, expand=1)
        
        scrollbar = tk.Scrollbar(frame, orient=tk.VERTICAL)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        sn_listbox = tk.Listbox(frame, selectmode=tk.MULTIPLE, yscrollcommand=scrollbar.set, font=("Arial", 12), bg="#333", fg="white")
        sn_listbox.pack(fill=tk.BOTH, expand=1)
        
        scrollbar.config(command=sn_listbox.yview)
        
        self.db_manager.cursor.execute("SELECT * FROM lavagem_placas WHERE serial = 'SN'")
        sn_records = self.db_manager.cursor.fetchall()

        for record in sn_records:
            sn_listbox.insert(tk.END, f"Linha Solicitante: {record[6]}, ID: {record[7]}, Serial: {record[8]}, Data: {record[1]}, Turno: {record[2]}, Hora: {record[3]}, Responsavel: {record[5]}")

        def delete_selected_sn():
            selected_indices = sn_listbox.curselection()
            selected_records = [sn_records[i] for i in selected_indices]
            
            justification = simpledialog.askstring("Justificativa", "Insira a justificativa para exclusão:")
            if justification:
                for record in selected_records:
                    self.db_manager.cursor.execute("DELETE FROM lavagem_placas WHERE id = ?", (record[0],))
                    self.db_manager.conn.commit()
                    
                    # Inserir justificativa no BDExclude
                    conn = sqlite3.connect('BDExclude.db')
                    cursor = conn.cursor()
                    cursor.execute("INSERT INTO BDExclude (Justificativa, serial) VALUES (?, ?)", (justification, 'SN'))
                    conn.commit()
                    conn.close()
                    
                sn_window.destroy()
                messagebox.showinfo("Registros Excluídos", f"{len(selected_records)} registros SN foram excluídos.")
            else:
                messagebox.showwarning("Justificativa Necessária", "Por favor, insira uma justificativa para a exclusão.")
        
        delete_button = tk.Button(sn_window, text="Excluir Selecionados", command=delete_selected_sn)
        delete_button.pack()


    def close(self):
        confirmation = messagebox.askyesno("Sair do Sistema", "Tem certeza que deseja sair?")
        if confirmation:
            self.db_manager.close()
            self.root.destroy()


def main():
    root = tk.Tk()
    window = DeleteRecordWindow(root)
    root.protocol("WM_DELETE_WINDOW", window.close)
    root.mainloop()

if __name__ == "__main__":
    main()
