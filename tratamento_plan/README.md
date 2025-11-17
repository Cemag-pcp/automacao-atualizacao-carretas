# Documentação do Arquivo `tratamento_plan.py`

## Objetivo
Este arquivo é responsável por tratar e organizar dados de planilhas relacionadas a carretas, produtos e processos, preparando um DataFrame final pronto para análise ou exportação. Ele realiza limpeza de dados, classificação, preenchimento de informações faltantes e integração com uma planilha de referência de células.

---

## Funcionalidades Principais

### 1. Limpeza e Normalização de Dados
- `limpar_espacos_unicode`: Remove pontos do início de strings e espaços Unicode extras.
- `contar_pontos_inicio`: Conta a quantidade de pontos no início de uma descrição.
- Normaliza nomes e códigos para facilitar o processamento.

### 2. Classificação e Definição de Processos
- `definir_primeiro_processo`: Identifica o primeiro processo do item (ex: MONTAR, PINTAR, EXPEDIR) com base na descrição da linha seguinte.
- `classificar_codigo`: Classifica o código do item em categorias como COMPONENTES, SECUNDÁRIOS, COMP. NAVAL.
- `observacao_diferente`: Define matéria-prima quando a observação não corresponde a processos conhecidos.

### 3. Identificação de Conjuntos e Produtos
- `buscar_conjuntos`: Associa cada item ao seu conjunto anterior com base no número de pontos.
- `verifica_descricao_3_depois`: Verifica se a descrição três linhas abaixo indica um processo especial.
- `definir_produto`: Determina o produto principal associado a cada linha, usando preenchimento progressivo (ffill).

### 4. Cálculo de Peso e Observações
- `definir_peso`: Define o peso de um item se a observação não for um processo padrão.
- `classificar_produto`: Classifica itens específicos como LATERAL, DIANTEIRA, TRASEIRA, 5ª RODA, CILINDRO, SUPORTE, ou mantém a célula original.

### 5. Extração e Concatenção de Informações
- `extrair_carreta`: Extrai o identificador da carreta a partir da descrição do produto, removendo sufixos.
- `concatenar_colunas`: Concatena múltiplas colunas em uma nova coluna, útil para gerar combinações como "peça + carreta".

### 6. Filtragem e Preparação Final
- Filtra apenas linhas com códigos e processos relevantes.
- Remove códigos que não devem ser incluídos na base final.
- Cria colunas auxiliares (`CELULA 1`, `CELULA 2`, `CELULA 3`) integrando dados da planilha de células.
- Preenche informações de carreta e produto para cada linha.
- Gera a coluna combinada `"peça + carreta"`.

### 7. Exportação
- O DataFrame final, contendo todas as colunas tratadas e classificadas, é exportado para um arquivo Excel (`carretas_final.xlsx`).
- A planilha final contém informações consolidadas sobre células, códigos, descrições, processos, produtos, conjuntos, peso e carretas.

---

## Regras de Negócio
1. Linhas com número de pontos zero definem novos produtos; as seguintes herdam o produto anterior.
2. Pontos iniciais em descrições são usados para identificar hierarquia de conjuntos.
3. Apenas processos e códigos relevantes são mantidos na base final.
4. Observações específicas influenciam o cálculo de matéria-prima e peso.
5. Conjuntos e produtos são tratados de forma a garantir consistência entre descrição, célula e carreta.
6. Todas as informações são integradas com uma planilha externa de referência de células para garantir padronização.

---

## Benefício
Permite transformar dados brutos de planilhas de carretas e processos em uma base limpa, consistente e pronta para análise ou importação em outros sistemas, garantindo que códigos, produtos, processos e carretas estejam corretamente classificados e vinculados.
