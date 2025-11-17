# Documentação do Arquivo `service.py`

## Objetivo
Este arquivo é responsável por criar uma aplicação web com Flask que fornece uma interface para execução manual e automática de um processo definido na função `main()` do módulo `api.py`. Além disso, gerencia o agendamento diário da execução do processo em segundo plano.

---

## Funcionalidades Principais

### 1. Aplicação Web
- Cria um servidor Flask (`app = Flask(__name__)`) para fornecer endpoints HTTP.
- Endpoints principais:
  - `"/"`: Renderiza a página `home.html`, funcionando como a interface inicial.
  - `"/executar/"` (POST): Executa manualmente a função `main()` do módulo `api.py`.
    - Retorna JSON com `{'status': 'ok'}` em caso de sucesso.
    - Retorna JSON com `{'status': 'erro'}` em caso de falha, além de registrar o erro no console.

### 2. Agendamento Automático
- A função `agendar_atualizacao()` agenda a execução diária da função `main()` às 01:00.
- Utiliza a biblioteca `schedule` para criar o job diário.
- Um loop contínuo verifica jobs pendentes a cada 30 segundos e os executa.
- Após as 01:10, imprime a lista de jobs já executados para controle/log.

### 3. Threads para Execução em Segundo Plano
- Para que o agendamento não bloqueie o servidor Flask, ele é executado em uma **thread separada**.
- Funções relacionadas:
  - `start_thread()`: Cria e inicia a thread daemon que executa o agendamento.
  - `init_agendamento()`: Inicializa o agendamento assim que o app Flask é configurado, garantindo que o contexto da aplicação esteja disponível.

### 4. Execução da Aplicação
- Quando o arquivo é executado diretamente (`if __name__ == '__main__':`), o servidor Flask é iniciado.
- A thread de agendamento é iniciada automaticamente antes da execução do servidor.

---

## Regras de Negócio
1. O endpoint `/executar/` permite disparar a execução do processo manualmente, retornando status de sucesso ou erro.
2. A execução automática ocorre todos os dias às 01:00, com verificação contínua de jobs pendentes a cada 30 segundos.
3. A execução agendada roda em segundo plano, de modo que o servidor Flask continue respondendo a requisições normalmente.
4. Qualquer exceção durante a execução manual ou agendada é capturada e registrada no console.

---

## Benefício
Permite tanto execução manual quanto automática de processos definidos, garantindo que tarefas críticas ocorram diariamente sem intervenção humana, enquanto mantém uma interface web simples para monitoramento ou acionamento manual.
