# 📋 PMPF Updater

<p align="center">
  <img src="https://img.shields.io/badge/Python-3.11+-3776AB?style=for-the-badge&logo=python&logoColor=white">
  <img src="https://img.shields.io/badge/Streamlit-FF4B4B?style=for-the-badge&logo=streamlit&logoColor=white">
  <img src="https://img.shields.io/badge/Pandas-150458?style=for-the-badge&logo=pandas&logoColor=white">
  <img src="https://img.shields.io/badge/openpyxl-217346?style=for-the-badge&logo=microsoftexcel&logoColor=white">
</p>

## 📌 Sobre

A **SEFA-PA** publica periodicamente portarias como as do **Preço Médio Ponderado ao Consumidor Final (PMPF)** de bebidas, utilizado como base de cálculo do ICMS por substituição tributária. Antes deste projeto, atualizar o catálogo de preços era um processo totalmente manual, um analista lia o PDF, identificava alterações e as inseria linha a linha no Excel.

O **PMPF Updater** automatiza esse processo: lê o PDF, compara com o catálogo atual e apresenta as mudanças detectadas numa interface web para revisão e aprovação do analista, estilo software de conciliação bancária.

---

## ⚙️ O que o sistema detecta e aplica

| Tipo | Descrição | O que aplica no Excel |
|---|---|---|
| 🔄 **Mudança de preço** | Mesmo GTIN com preço diferente | Nova linha em `PRECO-VIGENCIA` |
| 🆕 **Produto novo** | GTIN do PDF ausente no catálogo | Cadastro completo em `GTIN`, `PRODUTOS`, `CATALOGO` e `PRECO-VIGENCIA` |
| 🗑️ **Produto removido** | GTIN ativo no catálogo sumiu da portaria | `VÁLIDO = False` na aba `GTIN` |
| ✏️ **Descrição atualizada** | Nome diferente entre PDF e catálogo | Atualiza `MARCA/DESCRIÇÃO PORTARIA` + campo para o analista digitar o nome do catálogo |

> O arquivo original **nunca é modificado** — o sistema salva sempre uma cópia com sufixo de data/hora.

---

## 🗂️ Estrutura do projeto

```
pmpf_v3/
├── main.py          # Interface Streamlit (upload, revisão, aprovação)
├── extractor.py     # Extração de tabelas do PDF com Camelot
├── comparator.py    # Comparação PDF x catálogo
├── applier.py       # Gravação das mudanças aprovadas no Excel
├── normalizer.py    # Padronização de valores (MATERIAL, EMBALAGEM, etc.)
└── requirements.txt
```

---

## 🚀 Como executar

**Pré-requisitos:** Python 3.11+ e [Ghostscript](https://www.ghostscript.com/releases/gsdnld.html) instalado.

```bash
git clone https://github.com/victorpedrosa-ds/pmpf_v3.git
cd pmpf_v3
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
python -m streamlit run main.py
```

**Como usar:**
1. Fazer o upload do PDF da portaria e do catálogo Excel na barra lateral
2. Clique em **Processar**
3. Revise e aprove as mudanças nas 4 abas
4. Clique em **Aplicar e salvar** e baixe o Excel atualizado

**Obs:** Ainda estou trabalhando para melhorar o escalamento do código, tornando-o aplicável a diferentes tipos de documentos, com o objetivo de que possa ser utilizado por outras Secretarias de Fazenda ou por qualquer pessoa que lide com problemas semelhantes.

---

## 👤 Autor

**Victor Pedrosa**

[![LinkedIn](https://img.shields.io/badge/LinkedIn-0077B5?style=for-the-badge&logo=linkedin&logoColor=white)](https://linkedin.com/in/victorpedrosa-ds)
[![GitHub](https://img.shields.io/badge/GitHub-100000?style=for-the-badge&logo=github&logoColor=white)](https://github.com/victorpedrosa-ds)
