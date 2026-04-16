import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
import os

# ==========================================
# 1. FUNÇÃO: SPED (TXT) PARA EXCEL
# ==========================================
def exportar_sped_para_excel():
    arquivo_sped = filedialog.askopenfilename(
        title="1. Selecione o arquivo do SPED (TXT)",
        filetypes=[("Arquivos de Texto", "*.txt"), ("Todos os Arquivos", "*.*")]
    )
    if not arquivo_sped:
        return

    arquivo_excel = filedialog.asksaveasfilename(
        title="2. Onde deseja salvar a planilha formatada?",
        defaultextension=".xlsx",
        filetypes=[("Planilha do Excel", "*.xlsx")],
        initialfile="SPED_Formatado.xlsx"
    )
    if not arquivo_excel:
        return

    # Atualiza a interface para avisar que está trabalhando
    btn_excel.config(state=tk.DISABLED)
    btn_sped.config(state=tk.DISABLED)
    lbl_status.config(text="⏳ Gerando Excel... Isso pode levar vários minutos\ndependendo do tamanho do SPED.", fg="blue")
    root.update() # Força a tela a mostrar a mensagem antes de iniciar o trabalho pesado

    registros = {}
    
    try:
        with open(arquivo_sped, 'r', encoding='latin-1') as f:
            for numero_linha, linha in enumerate(f, start=1):
                linha = linha.strip()
                if not linha.startswith('|'):
                    continue
                
                campos = linha.split('|')[1:-1]
                if not campos:
                    continue
                    
                tipo_registro = campos[0]
                campos.insert(0, numero_linha)
                
                if tipo_registro not in registros:
                    registros[tipo_registro] = []
                    
                registros[tipo_registro].append(campos)

        with pd.ExcelWriter(arquivo_excel, engine='xlsxwriter') as writer:
            for reg, dados in registros.items():
                df = pd.DataFrame(dados)
                
                qtd_colunas = df.shape[1]
                nomes_colunas = ['Linha_Original', 'Registro'] + [f'Campo_{i+2:02d}' for i in range(qtd_colunas - 2)]
                df.columns = nomes_colunas
                
                nome_aba = f'Reg_{reg}'
                df.to_excel(writer, sheet_name=nome_aba, index=False)
                
                worksheet = writer.sheets[nome_aba]
                for idx, col in enumerate(df.columns):
                    tamanho_maximo = max(df[col].astype(str).map(len).max(), len(col)) + 2
                    tamanho_maximo = min(tamanho_maximo, 50) 
                    worksheet.set_column(idx, idx, tamanho_maximo)

        lbl_status.config(text="✅ Sucesso! Planilha gerada.", fg="green")
        messagebox.showinfo("Sucesso", f"Arquivo Excel gerado com sucesso em:\n{arquivo_excel}")
        
    except Exception as e:
        lbl_status.config(text="❌ Erro no processamento.", fg="red")
        messagebox.showerror("Erro", f"Ocorreu um erro ao gerar o Excel:\n{str(e)}")
        
    finally:
        # Reativa os botões quando terminar (dando certo ou errado)
        btn_excel.config(state=tk.NORMAL)
        btn_sped.config(state=tk.NORMAL)


# ==========================================
# 2. FUNÇÃO: EXCEL PARA SPED (TXT)
# ==========================================
def exportar_excel_para_sped():
    arquivo_excel = filedialog.askopenfilename(
        title="1. Selecione a Planilha do SPED (Excel)",
        filetypes=[("Planilhas do Excel", "*.xlsx"), ("Todos os Arquivos", "*.*")]
    )
    if not arquivo_excel:
        return

    arquivo_sped_saida = filedialog.asksaveasfilename(
        title="2. Onde deseja salvar o novo arquivo do SPED?",
        defaultextension=".txt",
        filetypes=[("Arquivo de Texto", "*.txt")],
        initialfile="SPED_Editado.txt"
    )
    if not arquivo_sped_saida:
        return

    # Atualiza a interface
    btn_excel.config(state=tk.DISABLED)
    btn_sped.config(state=tk.DISABLED)
    lbl_status.config(text="⏳ Reconstruindo SPED... Aguarde.", fg="blue")
    root.update()

    linhas_para_ordenar = []
    
    try:
        todas_abas = pd.read_excel(arquivo_excel, sheet_name=None, dtype=str)
        
        for nome_aba, df in todas_abas.items():
            for _, row in df.iterrows():
                linha_orig = int(row['Linha_Original'])
                
                valores = []
                for col in df.columns[1:]: 
                    val = row[col]
                    if pd.isna(val) or str(val).lower() == 'nan':
                        valores.append("")
                    else:
                        valores.append(str(val).strip())
                
                linha_txt = "|" + "|".join(valores) + "|"
                linhas_para_ordenar.append((linha_orig, linha_txt))
                
        linhas_para_ordenar.sort(key=lambda x: x[0])

        with open(arquivo_sped_saida, 'w', encoding='latin-1') as f:
            for _, linha_texto in linhas_para_ordenar:
                f.write(linha_texto + '\n')

        lbl_status.config(text="✅ Sucesso! SPED recriado.", fg="green")
        messagebox.showinfo("Sucesso", f"Arquivo SPED salvo e ordenado com sucesso em:\n{arquivo_sped_saida}")
        
    except Exception as e:
        lbl_status.config(text="❌ Erro no processamento.", fg="red")
        messagebox.showerror("Erro", f"Ocorreu um erro ao gerar o SPED:\n{str(e)}")
        
    finally:
        btn_excel.config(state=tk.NORMAL)
        btn_sped.config(state=tk.NORMAL)


# ==========================================
# 3. INTERFACE GRÁFICA (MENU PRINCIPAL)
# ==========================================
# Variáveis globais para a interface
root = tk.Tk()
root.title("Gerenciador SPED Fiscal")
root.geometry("380x250") # Aumentei um pouco a janela para caber o texto de status
# root.attributes('-topmost', True) <- REMOVIDO PARA EVITAR TRAVAMENTOS

label = tk.Label(root, text="Escolha a operação desejada:", font=("Arial", 11, "bold"))
label.pack(pady=15)

btn_excel = tk.Button(root, text="🔄 1. Converter SPED TXT para Excel", 
                        command=exportar_sped_para_excel, width=35, height=2)
btn_excel.pack(pady=5)

btn_sped = tk.Button(root, text="📝 2. Converter Excel para SPED TXT", 
                        command=exportar_excel_para_sped, width=35, height=2)
btn_sped.pack(pady=5)

# Novo rótulo de status na parte inferior da janela
lbl_status = tk.Label(root, text="Aguardando ação...", font=("Arial", 9), fg="gray")
lbl_status.pack(pady=15)

if __name__ == "__main__":
    root.mainloop()
