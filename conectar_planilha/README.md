# Documentação do Arquivo de Manipulação de Planilhas de Carretas (conextar_planilha.py)

## Objetivo
Este arquivo é responsável por gerenciar dados de planilhas relacionadas a carretas, incluindo a extração de informações, atualização de bases e integração de diferentes abas dentro de uma mesma planilha.

---

## Funcionalidades

### 1. Configuração e Conexão
- Autentica e conecta à API do Google Sheets usando um arquivo de credenciais.
- Retorna um cliente autorizado para manipular planilhas.

### 2. Extração de Apontamentos
- Conecta à planilha de apontamento e extrai as colunas `Código` e `Célula`.
- Filtra dados inválidos (códigos ou células vazias).
- Garante que apenas o último registro de cada código seja mantido.
- Retorna uma tabela com códigos e células atualizadas.

### 3. Armazenamento de Base Atualizada
- Substitui completamente os dados da aba `BASE ATUALIZADA` de uma planilha central com os dados fornecidos.
- Garante que a planilha reflita fielmente os dados do DataFrame.

### 4. Inserção Incremental de Carretas
- Adiciona dados ao final da aba `BASE ATUALIZADA` sem sobrescrever informações existentes.
- Preenche valores nulos com strings vazias.
- Mantém a integridade da planilha durante a inserção.

### 5. Integração de Peças
- Puxa os dados da aba `BASE PEÇAS` e adiciona ao final da aba `BASE ATUALIZADA`.
- Converte os dados da aba em DataFrame, preenche valores nulos e insere incrementalmente.
- Garante que informações de peças não sobrescrevam dados já existentes.

### 6. Consulta de Carretas da Base Externa
- Conecta a uma planilha externa (`base_felipe`) e extrai a lista de carretas únicas.
- Remove duplicidades e espaços em branco.
- Retorna a lista de carretas e informa a quantidade total.

---

## Regras de Negócio Gerais
1. Linhas sem código ou célula não são consideradas.
2. Atualizações podem ser totais (sobrescrevendo dados) ou incrementais (adicionando ao final).
3. Todos os dados são validados antes de serem inseridos ou atualizados.
4. Mantém consistência e integridade das planilhas durante as operações.

---

## Benefício
Permite centralizar, organizar e atualizar informações de carretas e peças em planilhas de forma consistente, garantindo que dados importantes estejam sempre disponíveis e corretos.
