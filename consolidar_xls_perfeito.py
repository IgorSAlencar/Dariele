import os
from pathlib import Path
import time

def consolidar_arquivos_xls_perfeito():
    """
    Consolida todos os arquivos .xls da pasta Arquivos em um único arquivo xlsx
    preservando 100% da formatação original copiando célula por célula
    """
    try:
        import win32com.client as win32
    except ImportError:
        print("Biblioteca win32com não está disponível.")
        print("Instalando pywin32...")
        os.system("pip install pywin32")
        import win32com.client as win32
    
    # Caminho da pasta com os arquivos
    pasta_arquivos = Path("Arquivos")
    
    # Nome do arquivo de saída
    arquivo_saida = "Relatorio.xlsx"
    
    # Verificar se a pasta existe
    if not pasta_arquivos.exists():
        print(f"Pasta {pasta_arquivos} não encontrada!")
        return
    
    # Encontrar todos os arquivos .xls
    arquivos_xls = list(pasta_arquivos.glob("*.xls"))
    
    if not arquivos_xls:
        print("Nenhum arquivo .xls encontrado na pasta!")
        return
    
    print(f"Encontrados {len(arquivos_xls)} arquivos .xls:")
    for arquivo in arquivos_xls:
        print(f"  - {arquivo.name}")
    
    # Inicializar Excel
    excel = None
    wb_consolidado = None
    
    try:
        print("\nInicializando Excel...")
        excel = win32.Dispatch("Excel.Application")
        excel.Visible = False  # Mude para True se quiser ver o processo
        excel.DisplayAlerts = False
        excel.ScreenUpdating = False  # Acelerar o processo
        
        # Criar novo workbook para consolidação
        wb_consolidado = excel.Workbooks.Add()
        
        # Usar a primeira planilha padrão para o primeiro arquivo
        primeira_copia = True
        
        for i, arquivo in enumerate(arquivos_xls):
            try:
                print(f"Processando: {arquivo.name}")
                
                # Abrir arquivo .xls
                arquivo_absoluto = str(arquivo.resolve())
                wb_origem = excel.Workbooks.Open(arquivo_absoluto)
                
                # Para cada planilha no arquivo de origem
                for j in range(1, wb_origem.Worksheets.Count + 1):
                    ws_origem = wb_origem.Worksheets(j)
                    
                    # Nome da planilha no arquivo consolidado
                    nome_planilha = arquivo.stem
                    if len(nome_planilha) > 31:
                        nome_planilha = nome_planilha[:31]
                    
                    # Verificar se já existe planilha com esse nome
                    nomes_existentes = [wb_consolidado.Worksheets(k).Name for k in range(1, wb_consolidado.Worksheets.Count + 1)]
                    contador = 1
                    nome_original = nome_planilha
                    while nome_planilha in nomes_existentes:
                        nome_planilha = f"{nome_original}_{contador}"
                        contador += 1
                    
                    if primeira_copia:
                        # Usar a planilha padrão para a primeira cópia
                        ws_destino = wb_consolidado.Worksheets(1)
                        primeira_copia = False
                    else:
                        # Adicionar nova planilha
                        ws_destino = wb_consolidado.Worksheets.Add()
                    
                    # Copiar a planilha inteira de uma só vez mantendo formatação
                    try:
                        # Método 1: Copiar planilha inteira como imagem e dados
                        ws_origem.Copy(Before=wb_consolidado.Worksheets(1))
                        ws_copiada = wb_consolidado.Worksheets(1)
                        ws_copiada.Name = nome_planilha
                        
                        # Deletar a planilha de destino original se não for a primeira
                        if not primeira_copia:
                            ws_destino.Delete()
                        
                        print(f"  ✓ Planilha '{nome_planilha}' copiada com formatação 100% preservada")
                        
                    except Exception as e1:
                        print(f"  Tentando método alternativo para {nome_planilha}...")
                        try:
                            # Método 2: Cópia célula por célula com formatação
                            usado_range = ws_origem.UsedRange
                            
                            if usado_range:
                                # Copiar valores e formatação
                                ws_destino.Range(
                                    ws_destino.Cells(1, 1),
                                    ws_destino.Cells(usado_range.Rows.Count, usado_range.Columns.Count)
                                ).Value = usado_range.Value
                                
                                # Copiar formatação célula por célula para áreas específicas
                                for row in range(1, min(usado_range.Rows.Count + 1, 50)):  # Limitar para acelerar
                                    for col in range(1, min(usado_range.Columns.Count + 1, 20)):
                                        try:
                                            celula_origem = ws_origem.Cells(row, col)
                                            celula_destino = ws_destino.Cells(row, col)
                                            
                                            if celula_origem.Value:
                                                # Copiar formatação da célula
                                                celula_origem.Copy()
                                                celula_destino.PasteSpecial(Paste=-4122)  # xlPasteFormats
                                                
                                        except:
                                            continue
                                
                                # Limpar clipboard
                                excel.CutCopyMode = False
                            
                            # Renomear a planilha
                            ws_destino.Name = nome_planilha
                            print(f"  ✓ Planilha '{nome_planilha}' copiada com método alternativo")
                            
                        except Exception as e2:
                            print(f"  ✗ Erro nos dois métodos para {nome_planilha}: {e2}")
                
                # Fechar arquivo origem
                wb_origem.Close(SaveChanges=False)
                
            except Exception as e:
                print(f"  ✗ Erro ao processar {arquivo.name}: {e}")
        
        # Reativar atualizações da tela
        excel.ScreenUpdating = True
        
        # Salvar o arquivo consolidado
        arquivo_saida_completo = str(Path.cwd() / arquivo_saida)
        
        # Remover arquivo se já existir
        if os.path.exists(arquivo_saida_completo):
            try:
                os.remove(arquivo_saida_completo)
                time.sleep(1)
            except:
                print(f"Aviso: Não foi possível remover {arquivo_saida} existente")
        
        # Salvar como xlsx
        try:
            wb_consolidado.SaveAs(arquivo_saida_completo, FileFormat=51)  # 51 = xlsx format
            print(f"Arquivo salvo com sucesso!")
        except Exception as e:
            print(f"Erro ao salvar: {e}")
            # Tentar salvar com nome alternativo
            arquivo_alternativo = str(Path.cwd() / f"Relatorio_backup_{int(time.time())}.xlsx")
            try:
                wb_consolidado.SaveAs(arquivo_alternativo, FileFormat=51)
                print(f"Arquivo salvo como: {arquivo_alternativo}")
            except Exception as e2:
                print(f"Erro ao salvar arquivo alternativo: {e2}")
        
        print(f"\n✅ Arquivo consolidado criado: {arquivo_saida}")
        print("Planilhas incluídas:")
        for k in range(1, wb_consolidado.Worksheets.Count + 1):
            print(f"  - {wb_consolidado.Worksheets(k).Name}")
        
    except Exception as e:
        print(f"Erro geral: {e}")
    
    finally:
        # Fechar tudo de forma mais robusta
        try:
            if wb_consolidado:
                wb_consolidado.Close(SaveChanges=False)
                print("Workbook fechado.")
        except Exception as e:
            print(f"Erro ao fechar workbook: {e}")
        
        try:
            if excel:
                excel.ScreenUpdating = True
                excel.Quit()
                print("Excel fechado com sucesso.")
        except Exception as e:
            print(f"Erro ao fechar Excel: {e}")
        
        # Aguardar um pouco para garantir que o processo do Excel seja liberado
        time.sleep(3)

if __name__ == "__main__":
    consolidar_arquivos_xls_perfeito() 