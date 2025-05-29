# 🏥 DICOM Dose Extractor Suite

Conjunto de ferramentas Python para extração e análise de dados de dose de radiação de arquivos DICOM SR (Structured Report) de Tomografia Computadorizada.

## 📋 Visão Geral

Este projeto oferece três ferramentas complementares para processar relatórios de dose DICOM:

- **🚀 DICOMDoseExtractor.py** - Extração direta DICOM → Excel (recomendado)
- **📊 DICOMDoseJSON.py** - Extração DICOM → JSON consolidado  
- **📈 DICOMDoseExcel.py** - Conversão JSON → Excel

## 🎯 Funcionalidades

### ✅ Extração Automática
- Busca recursiva por arquivos DICOM SR em estruturas de pastas complexas
- Identificação automática de arquivos DICOM válidos
- Extração de dados essenciais de dose de radiação CT

### 📊 Dados Extraídos
- **Informações do Paciente**: ID, Nome, Sexo, Data de Nascimento, Idade
- **Dados do Exame**: Data, Hospital, Equipamento
- **Parâmetros Técnicos**: kVp, mAs, Protocolo, Tipo de Aquisição
- **Métricas de Dose**: CTDIvol, DLP, DLP Total, SSDE, Tipo de Phantom

### 📈 Saídas Suportadas
- **Excel (.xlsx)**: Planilha formatada com dados organizados
- **JSON**: Arquivo estruturado para integração com outros sistemas

## 🛠️ Instalação

### Pré-requisitos
```bash
pip install pydicom openpyxl
```

### Dependências
- **Python 3.7+**
- **pydicom**: Leitura de arquivos DICOM
- **openpyxl**: Geração de planilhas Excel

## 🚀 Uso Rápido

### Opção 1: Extração Direta (Recomendada)
Para gerar diretamente a planilha Excel a partir dos DICOMs:

```bash
# Na pasta com subpastas contendo DICOMs
python DICOMDoseExtractor.py

# Ou especificando pasta e arquivo
python DICOMDoseExtractor.py --folder /caminho/para/dicoms --output relatorio_doses.xlsx
```

### Opção 2: Fluxo Completo (JSON + Excel)
Para quem precisa do JSON intermediário:

```bash
# 1. Gerar JSON consolidado
python DICOMDoseJSON.py

# 2. Converter JSON para Excel  
python DICOMDoseExcel.py dicom_reports_consolidated_TIMESTAMP.json
```

## 📖 Documentação Detalhada

### 🚀 DICOMDoseExtractor.py

**Extração direta DICOM → Excel (mais eficiente)**

```bash
# Uso básico
python DICOMDoseExtractor.py

# Opções avançadas
python DICOMDoseExtractor.py --folder /pasta/dicoms \
                            --output relatorio_2024.xlsx \
                            --debug
```

**Parâmetros:**
- `--folder, -f`: Pasta raiz para busca recursiva (padrão: pasta atual)
- `--output, -o`: Nome do arquivo Excel (padrão: ct_dose_direct_report.xlsx)
- `--debug, -d`: Ativa informações detalhadas de processamento

### 📊 DICOMDoseJSON.py

**Extração DICOM → JSON consolidado**

```bash
# Uso básico
python DICOMDoseJSON.py

# Opções avançadas
python DICOMDoseJSON.py --folder /pasta/dicoms \
                       --output dados_consolidados.json \
                       --debug
```

**Parâmetros:**
- `--folder, -f`: Pasta raiz para busca recursiva (padrão: pasta atual)
- `--output, -o`: Nome do arquivo JSON (padrão: dicom_reports_consolidated_TIMESTAMP.json)
- `--debug, -d`: Ativa informações detalhadas de processamento
- `--single, -s`: Processa um único arquivo DICOM específico

### 📈 DICOMDoseExcel.py

**Conversão JSON → Excel**

```bash
# Uso básico
python DICOMDoseExcel.py arquivo_dados.json

# Com arquivo de saída personalizado
python DICOMDoseExcel.py dados.json --output planilha_final.xlsx
```

**Parâmetros:**
- `json_file`: Arquivo JSON com dados DICOM (obrigatório)
- `--output, -o`: Nome do arquivo Excel (padrão: ct_dose_dicom_report.xlsx)

## 📁 Estrutura de Pastas Suportada

O projeto funciona com estruturas complexas de pastas, comuns em sistemas PACS:

```
pasta_principal/
├── DICOMDoseExtractor.py
├── 90d68f/
│   └── 82c48815/
│       └── 746ac786    # ← arquivo DICOM
├── a1b2c3/
│   └── d4e5f6/
│       └── g7h8i9      # ← arquivo DICOM
└── ... (milhares de outras pastas)
```

## 📊 Formato da Planilha Excel

A planilha gerada contém as seguintes colunas:

| Coluna | Descrição | Exemplo |
|--------|-----------|---------|
| ID do paciente | Identificador único do paciente | 12345 |
| Nome do paciente | Nome completo | João Silva |
| Sexo | M/F | M |
| Data de nascimento | Data formatada | Jan 15, 1980 |
| Idade | Calculada automaticamente | 44 |
| Pesquisa de interesse | Protocolo do exame | Chest Routine |
| Data do exame | Data e hora do exame | May 28, 2024, 14:30:15 |
| Descrição da série | Comentários da aquisição | Helical 1.25 |
| Scan mode | Tipo de aquisição | Helical |
| mAs | Corrente do tubo | 200 mA |
| kV | Tensão do tubo | 120 kV |
| CTDIvol | Índice de dose CT | 15.5 mGy |
| DLP | Produto dose-comprimento | 450.2 mGy*cm |
| DLP total | DLP acumulado do exame | 450.2 mGy*cm |
| Phantom type | Tipo de phantom | Body |
| SSDE | Dose específica por tamanho | 18.2 mGy |
| Avg scan size | Tamanho médio do scan | - |

## 🔧 Características Técnicas

### Códigos DICOM Suportados
- **113813**: Total DLP
- **113819**: CT Acquisition  
- **125203**: Acquisition Protocol
- **113830**: Mean CTDIvol
- **113734**: Tube Current
- **113733**: KVP
- **113838**: DLP
- **113930**: Size Specific Dose Estimation (SSDE)
- **113835**: Phantom Type

### Formatos de Data Suportados
- **DICOM padrão**: YYYYMMDD
- **Formatação de saída**: "May 28, 2024" ou "May 28, 2024, 14:30:15"
- **Cálculo de idade**: Automático baseado nas datas

### Tratamento de Dados
- **Valores numéricos**: ID do paciente e idade salvos como números no Excel
- **Valores ausentes**: Representados como "-" na planilha
- **Encoding**: UTF-8 para suporte a caracteres especiais
- **Validação**: Verificação automática de arquivos DICOM válidos

## 📈 Performance

### Otimizações Implementadas
- **Busca inteligente**: Verifica prefixo DICM antes da leitura completa
- **Leitura seletiva**: `stop_before_pixels=True` para arquivos grandes
- **Extração mínima**: Apenas campos necessários para o Excel
- **Memória eficiente**: Processamento arquivo por arquivo

### Capacidade
- ✅ **Milhares de arquivos**: Testado com grandes volumes
- ✅ **Estruturas complexas**: Navegação recursiva profunda
- ✅ **Arquivos grandes**: Otimizado para DICOMs de vários MB

## 🧪 Exemplos de Uso

### Cenário 1: Hospital com Sistema PACS
```bash
# Pasta do servidor PACS com milhares de subpastas
cd /servidor/pacs/estudos_ct_2024/
python DICOMDoseExtractor.py --output doses_hospital_2024.xlsx
```

### Cenário 2: Pesquisa Científica
```bash
# Gerar JSON para análise posterior + Excel para visualização
python DICOMDoseJSON.py --folder /pesquisa/dados_ct --debug
python DICOMDoseExcel.py dicom_reports_consolidated_*.json --output analise_doses.xlsx
```

### Cenário 3: Auditoria de Qualidade
```bash
# Processamento com informações detalhadas
python DICOMDoseExtractor.py --debug --output auditoria_$(date +%Y%m%d).xlsx
```

## 🔍 Troubleshooting

### Problemas Comuns

**❌ "Nenhum arquivo DICOM SR encontrado"**
- Verifique se os arquivos são realmente DICOM SR de dose
- Confirme a estrutura de pastas
- Use `--debug` para ver detalhes da busca

**❌ "Erro ao ler JSON"**
- Verifique se o arquivo JSON está válido
- Confirme o encoding (deve ser UTF-8)

**❌ "Erro ao salvar Excel"**
- Verifique permissões de escrita na pasta
- Confirme se o arquivo não está aberto em outro programa

### Logs e Debug
Use a opção `--debug` para informações detalhadas:
```bash
python DICOMDoseExtractor.py --debug
```

Isso mostrará:
- Arquivos encontrados e processados
- Erros específicos de cada arquivo
- Estatísticas de processamento

## 🤝 Contribuição

### Estrutura do Código
```
projeto/
├── DICOMDoseExtractor.py    # Extração direta → Excel
├── DICOMDoseJSON.py         # Extração → JSON
├── DICOMDoseExcel.py        # JSON → Excel
└── README.md                # Este arquivo
```

### Padrões de Código
- **Python 3.7+** compatível
- **Type hints** quando apropriado
- **Docstrings** para funções principais
- **Error handling** robusto

## 📄 Licença

Este projeto é fornecido como está, para uso em ambiente hospitalar e de pesquisa.

## 👥 Suporte

Para questões técnicas ou sugestões de melhorias, consulte a documentação ou entre em contato com a equipe de desenvolvimento.

---

**🏥 DICOM Dose Extractor Suite** - Facilitando a análise de doses de radiação em Tomografia Computadorizada