# Automacao_de_Indicadores_OnePage

Este projeto implementa uma solução de inteligência de dados para uma rede varejista com 25 unidades, automatizando o ciclo completo de ETL (Extração, Transformação e Carga) e a comunicação de resultados. O sistema processa bases de dados complexas para gerar relatórios de desempenho diários e anuais de forma totalmente autônoma.

Funcionalidades
Tratamento de Dados Multiformato: Utiliza a biblioteca Pandas para consolidar e tratar dados provenientes de arquivos Excel (.xlsx) e CSV, realizando o merge de tabelas de vendas, lojas e e-mails dos gestores.

Cálculo de Métricas de Negócio: Processa logicamente o faturamento total, a diversidade de produtos (mix de vendas) e o ticket médio, comparando os resultados do último dia disponível e do acumulado anual contra metas pré-estabelecidas.

Automação de Comunicação via Outlook: Integração robusta com o Microsoft Outlook através da biblioteca win32com.client para disparos de e-mails dinâmicos:

Relatórios OnePage: Envio de tabelas em HTML no corpo do e-mail com indicadores de status colorizados (verde/vermelho) e anexos personalizados por loja.

Relatórios Executivos: Consolidação de rankings de performance para a diretoria, destacando automaticamente as unidades de melhor e pior desempenho.

Gestão Estruturada de Arquivos: Uso da biblioteca pathlib para automação do sistema de arquivos, criando diretórios dinâmicos e organizando backups cronológicos de cada unidade.

Tecnologias Utilizadas
Python: Core da aplicação.

Pandas: Manipulação e análise de dados.

Pywin32 (win32com): Automação de interface com o Microsoft Outlook.

Pathlib: Manipulação inteligente de caminhos e diretórios.
