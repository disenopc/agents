# Official Company Website Finder

[![Python](https://img.shields.io/badge/python-3.11%2B-blue)](https://www.python.org/)  
[![License](https://img.shields.io/badge/license-MIT-green)](LICENSE)

This project allows you to automatically search for the official websites of companies using **Google Custom Search Engine (CSE)** and a scoring system to determine the most likely official URL.

---

## Features

- Generates optimized queries for Google CSE.
- Scores and filters results to identify the official website.
- Avoids URLs from social media and third-party sites.
- Saves results in Excel with search notes.
- Handles rate limits and connection errors with automatic pauses.

---

## Requirements

- Python 3.11+
- Required packages:

```bash

### Create environment from scratch

# Create the environment from scratch with Python 3.11
conda create -n prov python=3.11

# Activate the environment
conda activate prov

# Install useful packages
conda install -c conda-forge poetry

# Install OpenAI and LangChain if you plan to use LLM for classification
conda install -c conda-forge poetry 

# Install optional tools and generate environment.yml from the install history
conda env export --from-history > environment.yml
