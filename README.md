### Create environment from scratch

# Crear el entorno desde cero con Python 3.11
conda create -n prov python=3.11

# Activar el entorno
conda activate prov

# Instalar paquetes útiles
conda install -c conda-forge poetry

# Instalar OpenAI y LangChain si planeas usar LLM para clasificación
conda install -c conda-forge poetry 

# Instalar herramientas opcionales
conda env export --from-history > environment.yml
