# Sistema de Gerenciamento de Vendas

Este é um sistema de gerenciamento de vendas que permite adicionar, atualizar e remover produtos, gerar relatórios de vendas e enviar e-mails com os resultados.

## Pré-requisitos

* Python 3.x
* Bibliotecas Python: `csv`, `collections`, `plyer`, `tkinter`, `pandas`, `openpyxl`, `smtplib`, `email`, `matplotlib`, `seaborn`, `schedule`

Para instalar as bibliotecas necessárias, execute o seguinte comando:

```bash
pip install plyer pandas openpyxl matplotlib seaborn schedule

## Como Usar

1.  **Executar o programa:**

    Execute o script Python `Relatório de Vendas Python.py`.

2.  **Menu principal:**

    O programa exibirá um menu com as seguintes opções:

    * `1: Adicionar novo produto`
    * `2: Atualizar produto existente`
    * `3: Remover produto`
    * `4: Sair do programa`

    Digite o número da opção desejada e pressione Enter.

3.  **Adicionar novo produto:**

    * Digite o nome do produto e pressione Enter.
    * Digite a quantidade e pressione Enter.
    * Digite o preço e pressione Enter.
    * Repita o processo para adicionar mais produtos ou digite 'sair' para finalizar.

4.  **Atualizar produto existente:**

    * O programa exibirá a lista de produtos existentes.
    * Digite o nome do produto que deseja atualizar e pressione Enter.
    * Digite a nova quantidade e pressione Enter.
    * Digite o novo preço e pressione Enter.
    * Repita o processo para atualizar mais produtos ou digite 'sair' para finalizar.

5.  **Remover produto:**

    * O programa exibirá a lista de produtos existentes.
    * Digite o nome do produto que deseja remover e pressione Enter.
    * Repita o processo para remover mais produtos ou digite 'sair' para finalizar.

6.  **Relatório de vendas:**

    O programa gera um relatório de vendas com o total de vendas por produto, o produto mais vendido e o total geral de vendas. O relatório é exibido no console e salvo em um arquivo Excel (`relatorio_vendas.xlsx`). Um gráfico de barras com as vendas por produto também é gerado e salvo como `grafico_vendas.png`.

7.  **Enviar e-mail:**

    O programa envia um e-mail com o relatório de vendas em anexo. As credenciais de e-mail são solicitadas através de uma interface gráfica Tkinter.

    * **Criando uma senha de apps na Conta Google:**
        1.  Acesse sua Conta Google: [myaccount.google.com](https://myaccount.google.com/)
        2.  No painel de navegação à esquerda, clique em "Segurança".
        3.  Em "Como você faz login no Google", clique em "Verificação em duas etapas".
        4.  Siga as instruções para ativar a verificação em duas etapas.
        5.  Com a verificação em duas etapas ativada, volte para a seção "Segurança" da sua Conta Google.
        6.  Vá na barra de pesquisa e digite "Senhas de Apps", abra a opção que aparecer
        7.  Na tela que abrir digite o nome do app "Relatório de Vendas Python"
        8.  Clique em "Criar".
        9.  O Google exibirá uma senha de 16 caracteres. Anote essa senha para usar no aplicativo(você não consegue ver essa senha novamente).

8.  **Agendamento:**

    O programa está configurado para executar a análise diariamente às 9h.

##Arquivos

* `vendas.csv`: Armazena os dados de vendas.
* `relatorio_vendas.xlsx`: Armazena o relatório de vendas em formato Excel.
* `grafico_vendas.png`: Armazena o gráfico de vendas.

## Observações

* O programa utiliza notificações de desktop para exibir um resumo das vendas.
* Um pop-up com o resumo das vendas também é exibido.
* O programa trata erros para várias operações, como leitura de arquivos, envio de e-mail e exclusão de dados.
* O programa evita duplicatas no arquivo Excel.

## Melhorias futuras

* Interface gráfica para o menu principal.
* Melhorias na interface de e-mail.
* Refatoração de funções.
* Adição de logs.
* Documentação mais detalhada.
* Validação de dados.
* Organização de arquivos em pastas.