# 🗺️ Roteirizador e Gestor Inteligente de Carteiras

![Python](https://img.shields.io/badge/Python-3.9+-blue.svg)
![Streamlit](https://img.shields.io/badge/Streamlit-FF4B4B?style=flat&logo=Streamlit&logoColor=white)
![SQLite](https://img.shields.io/badge/SQLite-07405E?style=flat&logo=sqlite&logoColor=white)
![Google Maps API](https://img.shields.io/badge/Google%20Maps-API-4285F4?style=flat&logo=googlemaps&logoColor=white)

Um sistema corporativo completo em ambiente Web desenvolvido para otimização logística, roteirização inteligente e gestão visual de carteiras de clientes. Construído para suportar pesadas demandas de distribuição e análise de dados operacionais, o sistema utiliza algoritmos de cálculo de rotas (OSRM), geocodificação em massa e integração em tempo real com as APIs do Google Maps.

---

## ✨ Principais Funcionalidades

* **🚗 Roteirização com Balanceamento Dinâmico:** Criação de rotas otimizadas baseadas em distância (Haversine e OSRM), regras de negócio estruturadas por CEP e volume logístico. Permite distribuir clientes na semana priorizando entregas de maior volume no início ou fim do ciclo.
* **📍 Geocodificação Híbrida e Inteligente:** Converte endereços em coordenadas geográficas com fallback múltiplo (ArcGIS) e memória de cache (Banco SQL `memoria_geocoding.db`) para evitar requisições repetidas, reduzindo o tempo de processamento de minutos para segundos.
* **👁️ Dashboards Interativos (Google Maps):** Renderização de mapas táticos de alta performance. Exibição de pins personalizados por vendedor, traçado de rotas rua a rua e integração nativa com o **Google Street View** em popups interativos.
* **🚧 Radar de Correção (Machine Learning Passivo):** Identifica endereços não localizados e permite a correção manual. O sistema aprende e salva a coordenada correta no Banco de Dados para que o erro nunca mais se repita nas próximas importações.
* **🗄️ Histórico em Banco de Dados (SQLite):** Salvamento seguro das rotas geradas e sobras logísticas, permitindo o recarregamento instantâneo de cenários anteriores.
* **📦 Exportação de Resultados:** Geração automática de arquivos `.xlsx` roteirizados e download de pacotes ZIP consolidados com KPIs de performance.

---

## 🛠️ Tecnologias Utilizadas

* **Linguagem:** Python
* **Interface Web (Frontend):** Streamlit
* **Manipulação e Análise de Dados:** Pandas, Numpy
* **Geolocalização & Roteirização:** Geopy (ArcGIS), Requests (OSRM API)
* **Visualização Espacial:** Google Maps JavaScript API, Google Street View Static API, Injeção nativa de HTML/CSS/JS
* **Banco de Dados:** SQLite3

---

## 🚀 Como Executar o Projeto Localmente

### 1. Pré-requisitos
* Ter o Python instalado (versão 3.9 ou superior).
* Possuir uma chave ativa do Google Cloud com as seguintes APIs habilitadas e com conta de faturamento (Billing) ativa:
  * `Maps JavaScript API`
  * `Street View Static API`

### 2. Instalação
Clone este repositório em sua máquina local:
```bash
git clone [https://github.com/seu-usuario/roteirizador-inteligente.git](https://github.com/seu-usuario/roteirizador-inteligente.git)
cd roteirizador-inteligente
```

python -m venv venv
# No Windows:
venv\Scripts\activate
# No Linux/Mac:
source venv/bin/activate

GOOGLE_MAPS_API_KEY = "SUA_CHAVE_AQUI"

streamlit run V9_GESTOR_CARTEIRA_ROTAS.py

📖 Guia de Uso Completo (Manual da Aplicação)
O sistema é dividido em módulos principais, acessíveis através da barra lateral de navegação:

1. Aba Roteirizador: Faça o upload da base do ERP (Excel ou CSV). O sistema fará o mapeamento das colunas automaticamente. Defina os dias de trabalho no calendário, o volume mínimo/máximo e clique em processar. Corrija alertas de endereço no "Radar" e baixe a rota pronta.

2. Visualizador Operacional: Suba planilhas que já possuem datas e rotas pré-definidas para ver o desenho tático no mapa, avaliar o balanço semanal e visualizar o trajeto no mapa de calor.

3. Histórico (SQL): Acesse todas as rotas processadas no mês, recarregue dados antigos diretamente da memória ou apague rotas antigas.

⚠️ Solução de Problemas Comuns (Troubleshooting)
Mapa não carrega / Tela Escura ("For development purposes only"): Verifique se o seu projeto no Google Cloud possui uma Conta de Faturamento (Billing) ativa. O cartão de crédito é obrigatório para liberar o mapa web.

Foto da Fachada Quebrada: Certifique-se de que a API Street View Static API está ativada e sem restrições na sua chave.

Sistema Lento na Primeira Vez: A primeira importação de uma base grande pode demorar devido à busca no satélite. Nas importações seguintes da mesma base, a leitura levará apenas segundos devido à arquitetura de Cache em SQLite implantada no sistema.

💡 Contribuições
Contribuições, relatórios de bugs e solicitações de novos recursos são bem-vindos! Sinta-se à vontade para abrir uma issue ou enviar um pull request.

📄 Licença
Desenvolvido para modernizar a logística corporativa, análise espacial e distribuição inteligente.
