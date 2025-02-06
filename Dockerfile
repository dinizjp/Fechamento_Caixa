# Use uma imagem oficial do Python (por exemplo, 3.9-slim)
FROM python:3.9-slim

# Instale dependências do sistema necessárias (curl, gnupg, apt-transport-https e unixodbc-dev)
RUN apt-get update && apt-get install -y \
    curl \
    gnupg \
    apt-transport-https \
    unixodbc-dev \
    && rm -rf /var/lib/apt/lists/*

# Adicione o repositório da Microsoft para o driver ODBC
RUN curl https://packages.microsoft.com/keys/microsoft.asc | apt-key add - \
    && curl https://packages.microsoft.com/config/ubuntu/20.04/prod.list > /etc/apt/sources.list.d/mssql-release.list

# Atualize os repositórios e instale o msodbcsql18 (aceitando o EULA)
RUN apt-get update && ACCEPT_EULA=Y apt-get install -y msodbcsql18

# Defina o diretório de trabalho
WORKDIR /app

# Copie o arquivo requirements.txt para o container e instale as dependências Python
COPY requirements.txt .
RUN pip install --upgrade pip && pip install -r requirements.txt

# Copie todo o restante do código para o container
COPY . .

# Exponha a porta que o Streamlit usa (padrão 8501)
EXPOSE 8501

# Comando para rodar a aplicação
CMD ["streamlit", "run", "app.py", "--server.port=8501", "--server.address=0.0.0.0"]
