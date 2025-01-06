import base64
import os
import openpyxl

def processar_planilha(caminho_planilha, aba, coluna_codigo, coluna_nome_produto, coluna_caminho_imagem):
    # Abrir a planilha
    workbook = openpyxl.load_workbook(caminho_planilha)
    sheet = workbook[aba]

    # Caminho onde as imagens serão salvas
    caminho_pasta = "C:/site/ft"
    os.makedirs(caminho_pasta, exist_ok=True)  # Cria a pasta se ela não existir

    # Iterar pelas linhas da planilha, começando da segunda linha (ignorando o cabeçalho)
    for idx, row in enumerate(sheet.iter_rows(min_row=2, values_only=True), start=2):
        codigo_base64 = row[coluna_codigo - 1]
        nome_produto = row[coluna_nome_produto - 1]

        # Verificar se o código Base64 está em branco ou nulo
        if not codigo_base64:
            print(f"Pulando linha {idx} pois o código Base64 está em branco.")
            continue

        # Verificar se o comprimento do código Base64 é um múltiplo de 4
        while len(codigo_base64) % 4 != 0:
            codigo_base64 += '='  # Adicionar padding

        # Decodificar o código Base64
        try:
            imagem_decodificada = base64.b64decode(codigo_base64)
        except Exception as e:
            print(f"Erro ao decodificar o código Base64 na linha {idx}: {e}")
            continue

        # Caminho completo para salvar a imagem
        caminho_imagem = os.path.join(caminho_pasta, f"{nome_produto}.png")

        # Salvar a imagem como PNG
        with open(caminho_imagem, "wb") as imagem_arquivo:
            imagem_arquivo.write(imagem_decodificada)

        # Escrever o caminho da imagem na coluna M (coluna 13)
        sheet.cell(row=idx, column=coluna_caminho_imagem, value=f"=HYPERLINK(\"file://{caminho_imagem}\", \"Link\")")

        print(f"Imagem para {nome_produto} salva em {caminho_imagem}")

    # Salvar as mudanças na planilha
    workbook.save(caminho_planilha)

# Configurações da planilha
caminho_planilha = "C:/site/dada.xlsx"  # Substitua pelo caminho correto
aba = "DADA"
coluna_codigo = 2  # Substitua pelo número da coluna que contém o código Base64
coluna_nome_produto = 12  # Substitua pelo número da coluna que contém o nome do produto
coluna_caminho_imagem = 13  # Coluna M

# Processar a planilha
processar_planilha(caminho_planilha, aba, coluna_codigo, coluna_nome_produto, coluna_caminho_imagem)
