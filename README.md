📌 Funcionalidades
✅ Cadastro e controle de equipamentos
✅ Geração de relatórios fotográficos em PDF
✅ Inclusão automática de imagens no relatório
✅ QR Code com link da empresa
✅ Organização automática por modelo e contrato
✅ Filtro de relatórios por data e modelo
✅ Abertura rápida de relatórios gerados
🧠 Diferenciais do sistema
📄 Relatórios profissionais com layout personalizado
🖼️ Inserção automática de múltiplas imagens
🏷️ Nomeação inteligente de arquivos
📁 Organização automática em pastas
⚙️ Configuração de diretório base para relatórios
🛠️ Tecnologias utilizadas
Python
ReportLab
qrcode
CustomTkinter
Pandas
📂 Estrutura do projeto
📁 Projeto
 ┣ 📁 utils
 ┣ 📁 services
 ┣ 📁 telas
 ┣ 📁 assets
 ┣ 📁 dist
 ┣ 📁 build
 ┣ 📄 main.py
 ┣ 📄 README.md
 ┗ 📄 config_relatorios.json
📄 Geração de Relatórios

O sistema gera relatórios contendo:

Logo da empresa
Dados do equipamento
Imagens organizadas automaticamente
QR Code com link da empresa
Layout profissional com marca d’água
▶️ Como executar o projeto
1. Clone o repositório
git clone https://github.com/SEU_USUARIO/SEU_REPOSITORIO.git
2. Acesse a pasta
cd SEU_REPOSITORIO
3. Instale as dependências
pip install -r requirements.txt
4. Execute o sistema
python main.py
⚙️ Configuração

O sistema permite configurar a pasta base dos relatórios:

Arquivo: config_relatorios.json
Diretório criado automaticamente se não existir
📊 Organização dos arquivos

Os relatórios são organizados automaticamente:

/Relatorios/
   ┣ 📁 MODELO_X
   ┃ ┣ Relatorio_modelo_contrato_data.pdf
📅 Filtros disponíveis
🔎 Por modelo
📆 Por data (formato: dd/mm/aaaa)
📷 Imagens do sistema


📄 Licença

Uso interno / educacional.

👨‍💻 Autor

Desenvolvido por José Victor

⭐ Futuras melhorias
Dashboard com indicadores
Integração com banco de dados
Exportação para Excel
Sistema de usuários
