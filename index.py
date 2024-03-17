import time
import pandas as pd
from matplotlib import pyplot as plt
from io import BytesIO
import pyautogui as pyau
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Image
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.units import inch

# Mapeamento da tela
botaoArquivo = (130, 201)
botaoFazerDownload = (158, 567)
opcoesDownload = (788, 574)
baixarCSV = (749, 771)
barraPesquisa = (268, 92)
barraWin = (143, 143)

# URL da planilha do Google Sheets
url_google_sheets = 'https://docs.google.com/spreadsheets/d/1J-EPUqG5X4t2tlle7OusCuvBH5Nhl2IUAmMo2KgLHHQ/edit#gid=700435573'

# Abrir o Google Chrome
time.sleep(3)
pyau.hotkey('win')
pyau.moveTo(barraWin, duration=1)
pyau.click()
pyau.write('chrome')
time.sleep(1)
pyau.hotkey('enter')

# Verificar se a janela está maximizada
window_is_maximized = pyau.getActiveWindow().isMaximized

# Se a janela não estiver maximizada, maximizá-la
if not window_is_maximized:
    pyau.hotkey('win', 'up')  # Use o atalho do teclado do Windows para maximizar a janela

# Aguarde um momento para o Chrome se ajustar ao tamanho máximo
time.sleep(5)

# Mover o mouse para a barra de pesquisa e abrir a URL
pyau.moveTo(barraPesquisa, duration=1)
pyau.click()
pyau.write(url_google_sheets)
pyau.hotkey('enter')

# Esperar um momento para a página carregar
time.sleep(7)  # Aumentei o tempo de espera para 20 segundos

# Baixar a planilha CSV com os dados
pyau.moveTo(botaoArquivo, duration=1)
pyau.click()
pyau.moveTo(botaoFazerDownload, duration=1)
time.sleep(1)  # Esperar um momento antes de mover para as opções de download
pyau.moveTo(opcoesDownload, duration=1)
time.sleep(1)  # Esperar um momento antes de mover para o botão de download CSV
pyau.moveTo(baixarCSV, duration=1)
pyau.click()
time.sleep(5)

# Caminho para o arquivo CSV baixado
caminho_arquivo = r'C:\Users\GenBR114\Downloads\dados - dados.csv'

# Carregar os dados do arquivo CSV para um DataFrame do Pandas
dados = pd.read_csv(caminho_arquivo)

# Resumo Estatistico
resumo_estatistico = dados.describe()

# Tendencias Temporais
dados['Data da Transação'] = pd.to_datetime(dados['Data da Transação'])
transacoes_por_mes = dados.groupby(dados['Data da Transação'].dt.to_period('M')).size()

# Segmentação de Clientes
segmentos = dados.groupby('Identificador do Cliente').agg({'Valor da Transação': 'sum'}).sort_values(by='Valor da Transação', ascending=False)

# Análise de Categoria
analise_categoria = dados.groupby('Categoria da Transação').agg({'Valor da Transação': 'sum'})

# Função para obter as preferências dos clientes com base nos tipos de transações mais frequentes
def get_preferencia_cliente(cliente_id):
    transacoes_cliente = dados[dados['Identificador do Cliente'] == cliente_id]
    tipo_transacao_mais_frequente = transacoes_cliente['Categoria da Transação'].value_counts().idxmax()
    return tipo_transacao_mais_frequente


# Detecção de Anomalias
media = dados['Valor da Transação'].mean()
desvio_padrao = dados['Valor da Transação'].std()
limite_anomalias = media + 3 * desvio_padrao
transacoes_anomalas = dados[dados['Valor da Transação'] > limite_anomalias]

# Criar um PDF
pdf_file = "relatorio.pdf"
doc = SimpleDocTemplate(pdf_file, pagesize=letter)

# Estilos para os parágrafos
styles = getSampleStyleSheet()
titulo_estilo = styles["Title"]
subtitulo_estilo = styles["Heading2"]
texto_estilo = styles["Normal"]

# Lista para armazenar os elementos do PDF
elements = []

# Adicionar título
elements.append(Paragraph("Relatório Financeiro", titulo_estilo))
elements.append(Spacer(1, 12))

# Adicionar Resumo Estatístico
elements.append(Paragraph("Resumo Estatístico", subtitulo_estilo))
resumo_texto = f"Valor Médio da Transação: R${resumo_estatistico.loc['mean', 'Valor da Transação']:.2f}<br/>" \
               f"Maior Valor de Transação: R${resumo_estatistico.loc['max', 'Valor da Transação']:.2f}<br/>" \
               f"Número Total de Clientes: {len(segmentos)}"
elements.append(Paragraph(resumo_texto, texto_estilo))
elements.append(Spacer(1, 12))

# Adicionar Tendências Temporais das Transações
elements.append(Paragraph("Tendências Temporais das Transações", subtitulo_estilo))
plt.plot(transacoes_por_mes.index.strftime('%Y-%m'), transacoes_por_mes.values, marker='o')
plt.xlabel('Data da Transação (Ano-Mês)')
plt.ylabel('Número de Transações')
plt.title('Tendências Temporais das Transações')
plt_buffer = BytesIO()
plt.savefig(plt_buffer, format='png')
plt.close()
elements.append(Image(plt_buffer, width=6*inch, height=3*inch))
elements.append(Spacer(1, 12))

# Adicionar Segmentação de Clientes ao PDF
elements.append(Paragraph("Segmentação de Clientes", styles['Heading1']))
segmentos_text = "Top 5 Clientes por Valor Total de Transações:<br/>"
for i, (cliente_id, valor_transacao) in enumerate(segmentos.head().iterrows(), start=1):
    preferencia_cliente = get_preferencia_cliente(cliente_id)
    segmentos_text += f"{i}. Cliente {cliente_id}: R${valor_transacao['Valor da Transação']:.2f} - Preferência: {preferencia_cliente}<br/>"
elements.append(Paragraph(segmentos_text, styles['Normal']))
elements.append(Spacer(1, 12))  # Adicionar um espaçamento entre os clientes

# Adicionar Análise de Categoria
elements.append(Paragraph("Análise de Categoria", subtitulo_estilo))
plt.bar(analise_categoria.index, analise_categoria['Valor da Transação'])
plt.xlabel('Categoria da Transação')
plt.ylabel('Valor Total da Transação')
plt.title('Análise de Categoria')
plt_buffer_analise_categoria = BytesIO()
plt.savefig(plt_buffer_analise_categoria, format='png')
plt.close()
elements.append(Image(plt_buffer_analise_categoria, width=6*inch, height=3*inch))
elements.append(Spacer(1, 12))

# Adicionar Detecção de Anomalias
elements.append(Paragraph("Detecção de Anomalias", subtitulo_estilo))
if not transacoes_anomalas.empty:
    anomalias_texto = "Transações Anômalas Encontradas:<br/>"
    for _, transacao in transacoes_anomalas.iterrows():
        anomalias_texto += f"Transação em {transacao['Data da Transação']}: R${transacao['Valor da Transação']:.2f}<br/>"
else:
    anomalias_texto = "Nenhuma transação anômala encontrada."
elements.append(Paragraph(anomalias_texto, texto_estilo))
elements.append(Spacer(1, 12))

# Criar um PDF
pdf_file = "relatorio.pdf"
doc = SimpleDocTemplate(pdf_file, pagesize=letter)

# Adicionar os elementos ao PDF
doc.build(elements)

# Após a criação do relatório PDF
pyau.alert(f"Relatório gerado com sucesso em '{pdf_file}'", title="Relatório Gerado")
