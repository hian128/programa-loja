import pandas as pd
import tkinter as tk
from tkinter import messagebox, ttk

# Nome do arquivo
nome_arquivo = "Estoque.xlsx"

# ---------------------- Funções ----------------------
def ler_estoque():
    try:
        df = pd.read_excel(nome_arquivo, sheet_name='Estoque')
    except FileNotFoundError:
        df = pd.DataFrame(columns=["ID do produto", "Produto", "Marca",
                                   "Preço de custo", "Preço de venda", "Entrada"])
        df.to_excel(nome_arquivo, sheet_name='Estoque', index=False)
        df = pd.read_excel(nome_arquivo, sheet_name='Estoque')
    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao ler o arquivo: {e}")
        return None

    if "Saída" not in df.columns:
        df["Saída"] = 0
    if "Total estoque" not in df.columns:
        df["Total estoque"] = df["Entrada"] - df["Saída"]
    if "Lucro do dia" not in df.columns:
        df["Lucro do dia"] = 0

    df["Lucro do dia"] = (df["Preço de venda"] - df["Preço de custo"]) * df["Saída"]

    return df

def salvar_estoque(df):
   
    df.to_excel(nome_arquivo, sheet_name='Estoque', index=False)

# ---------------------- Tela de adicionar produto ----------------------
def abrir_formulario_produto():
    df = ler_estoque()
    if df is None:
        return

    form_window = tk.Toplevel(root)
    form_window.title("Adicionar Produto")
    form_window.geometry("400x400")
    form_window.configure(bg="#f0f0f0")

    # Labels e entradas
    tk.Label(form_window, text="Nome do Produto:", bg="#f0f0f0").pack(pady=5)
    nome_entry = tk.Entry(form_window, width=30)
    nome_entry.pack(pady=5)

    tk.Label(form_window, text="Marca:", bg="#f0f0f0").pack(pady=5)
    marca_entry = tk.Entry(form_window, width=30)
    marca_entry.pack(pady=5)

    tk.Label(form_window, text="Preço de compra:", bg="#f0f0f0").pack(pady=5)
    preco_compra_entry = tk.Entry(form_window, width=30)
    preco_compra_entry.pack(pady=5)

    tk.Label(form_window, text="Preço de venda:", bg="#f0f0f0").pack(pady=5)
    preco_venda_entry = tk.Entry(form_window, width=30)
    preco_venda_entry.pack(pady=5)

    tk.Label(form_window, text="Quantidade:", bg="#f0f0f0").pack(pady=5)
    qtd_entry = tk.Entry(form_window, width=30)
    qtd_entry.pack(pady=5)

    # Função para salvar os dados
    def salvar_produto():
        try:
            nome_produto = nome_entry.get()
            marca_produto = marca_entry.get()
            preco_compra = float(preco_compra_entry.get())
            preco_venda = float(preco_venda_entry.get())
            qtd_produto = int(qtd_entry.get())

            novo_produto = pd.DataFrame([{
                'ID do produto': len(df) + 1,
                'Produto': nome_produto,
                'Marca': marca_produto,
                'Preço de custo': preco_compra,
                'Preço de venda': preco_venda,
                'Entrada': qtd_produto,
                'Saída': 0,
                'Total estoque': qtd_produto
            }])
            if preco_compra > preco_venda:
                messagebox.showerror("Erro", "O preço de venda deve ser maior que o preço de custo.")
                return
            df_novo = pd.concat([df, novo_produto], ignore_index=True)
            salvar_estoque(df_novo)
            messagebox.showinfo("Sucesso", f"Produto '{nome_produto}' adicionado com sucesso!")
            form_window.destroy()

        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao adicionar produto: {e}")

    tk.Button(form_window, text="Adicionar Produto", bg="#457b9d", fg="white",
              font=("Helvetica", 12), width=20, command=salvar_produto).pack(pady=20)

# ---------------------- Tela de vender produto ----------------------
def abrir_formulario_venda():
    df = ler_estoque()
    if df is None:
        return

    form_window = tk.Toplevel(root)
    form_window.title("Vender Produto")
    form_window.geometry("400x300")
    form_window.configure(bg="#f0f0f0")

    tk.Label(form_window, text="Selecione o produto (ID):", bg="#f0f0f0").pack(pady=5)
    produto_var = tk.StringVar()
    produto_combo = ttk.Combobox(form_window, textvariable=produto_var, state="readonly", width=28)
    produto_combo['values'] = [f"{row['ID do produto']} - {row['Produto']}" for _, row in df.iterrows()]
    produto_combo.pack(pady=5)

    tk.Label(form_window, text="Quantidade a vender:", bg="#f0f0f0").pack(pady=5)
    qtd_entry = tk.Entry(form_window, width=30)
    qtd_entry.pack(pady=5)

    def registrar_venda():
        try:
            if not produto_var.get():
                messagebox.showerror("Erro", "Selecione um produto!")
                return

            entrada_id = int(produto_var.get().split(" - ")[0])
            qtd_venda = int(qtd_entry.get())

            idx = df.index[df["ID do produto"] == entrada_id][0]

            qtd_entrada = df.at[idx, "Entrada"]
            qtd_saida = df.at[idx, "Saída"] if not pd.isna(df.at[idx, "Saída"]) else 0
            total_estoque = qtd_entrada - qtd_saida

            if qtd_venda > total_estoque:
                messagebox.showerror("Erro", "Quantidade insuficiente no estoque!")
                return

            df.at[idx, "Saída"] = qtd_saida + qtd_venda
            df.at[idx, "Total estoque"] = qtd_entrada - df.at[idx, "Saída"]

            salvar_estoque(df)
            messagebox.showinfo("Sucesso", f"Venda de {qtd_venda} unidades registrada com sucesso!")
            form_window.destroy()

        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao vender produto: {e}")

    tk.Button(form_window, text="Registrar Venda", bg="#e63946", fg="white",
              font=("Helvetica", 12), width=20, command=registrar_venda).pack(pady=20)

def abrir_formulario_remover():
    df = ler_estoque()
    if df is None:
        return

    form_window = tk.Toplevel(root)
    form_window.title("Remover Produto")
    form_window.geometry("400x250")
    form_window.configure(bg="#f0f0f0")

    tk.Label(form_window, text="Selecione o produto (ID):", bg="#f0f0f0").pack(pady=10)
    
    produto_var = tk.StringVar()
    produto_combo = ttk.Combobox(form_window, textvariable=produto_var, state="readonly", width=35)
    produto_combo['values'] = [f"{row['ID do produto']} - {row['Produto']}" for _, row in df.iterrows()]
    produto_combo.pack(pady=5)

    def remover_produto():
        try:
            if not produto_var.get():
                messagebox.showerror("Erro", "Selecione um produto para remover!")
                return

            produto_id = int(produto_var.get().split(" - ")[0])

            # Localizar linha
            df_filtrado = df[df["ID do produto"] != produto_id]

            salvar_estoque(df_filtrado)

            messagebox.showinfo("Sucesso", f"Produto removido com sucesso!")
            form_window.destroy()

        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao remover produto: {e}")

    tk.Button(
        form_window, 
        text="Remover Produto", 
        bg="#e63946", 
        fg="white", 
        font=("Helvetica", 12),
        width=20, 
        command=remover_produto
    ).pack(pady=20)

# ---------------------- Mostrar estoque ----------------------
def mostrar_estoque():
    df = ler_estoque()
    if df is None:
        return

    estoque_window = tk.Toplevel(root)
    estoque_window.title("Estoque Atual")
    estoque_window.geometry("900x400")
    estoque_window.configure(bg="#f0f0f0")

    tree = ttk.Treeview(estoque_window)
    tree.pack(fill="both", expand=True)

    style = ttk.Style()
    style.theme_use("default")
    style.configure("Treeview",
                    background="#eeeeee",
                    foreground="black",
                    rowheight=25,
                    fieldbackground="#eeeeee")
    style.map("Treeview", background=[("selected", "#347083")])

    tree["columns"] = list(df.columns)
    tree["show"] = "headings"
    for col in df.columns:
        tree.heading(col, text=col)
        tree.column(col, width=100, anchor="center")

    for i, (_, row) in enumerate(df.iterrows()):
        if i % 2 == 0:
            tree.insert("", "end", values=list(row), tags=("evenrow",))
        else:
            tree.insert("", "end", values=list(row), tags=("oddrow",))

    tree.tag_configure("evenrow", background="#ffffff")
    tree.tag_configure("oddrow", background="#d3d3d3")

def calcular_lucro_do_dia():
    df = ler_estoque()
    if df is None:
        return

    # Garantir que a coluna exista
    if "Lucro do dia" not in df.columns:
        df["Lucro do dia"] = 0

    df["Lucro do dia"] = (df["Preço de venda"] - df["Preço de custo"]) * df["Saída"]

    salvar_estoque(df)
    messagebox.showinfo("Sucesso", "Lucro do dia calculado e salvo na planilha!")

def remover_produto(produto_id):
    df = ler_estoque()
    if df is None:
        return

    if produto_id in df["ID do produto"].values:
         
        df = df[df["ID do produto"] != produto_id]
        salvar_estoque(df)
        messagebox.showinfo("Sucesso", f"Produto com ID {produto_id} removido com sucesso!")
    else:
        messagebox.showerror("Erro", f"Produto com ID {produto_id} não encontrado!")

# ---------------------- Janela principal ----------------------
root = tk.Tk()
root.title("Sistema de Estoque")
root.geometry("450x350")
root.configure(bg="#a8dadc")
root.eval('tk::PlaceWindow . center')

tk.Label(root, text="Sistema de Estoque", font=("Helvetica", 20, "bold"),
         bg="#a8dadc", fg="#1d3557").pack(pady=20)

btn_style = {"font": ("Helvetica", 14), "width": 20, "bg": "#457b9d",
             "fg": "white", "activebackground": "#1d3557", "activeforeground": "white",
             "bd":0, "relief":"flat"}

tk.Button(root, text="Ver Estoque", command=mostrar_estoque, **btn_style).pack(pady=5)
tk.Button(root, text="Adicionar Produto", command=abrir_formulario_produto, **btn_style).pack(pady=5)
tk.Button(root, text="Vender Produto", command=abrir_formulario_venda, **btn_style).pack(pady=5)
tk.Button(root, text="Remover Produto", command=abrir_formulario_remover, **btn_style).pack(pady=5)
tk.Button(root, text="Sair", command=root.quit, **btn_style).pack(pady=20)


root.mainloop()

