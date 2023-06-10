# Email_Automatico
    Análise e Desenvolvimento de Dados e envio de email automático com Python.

### Objetivo e Cenário para esta aplicação em Python:

Imagine que você trabalha em uma grande rede de lojas com 25 unidades espalhadas por todo o Brasil.

Todo dia, pela manhã, o analista de dados calcula os Resumos (diários e anuais) e envia para o gerente de cada loja o Resumo da sua própria loja, bem como todas as informações usadas no cálculo dos indicadores. Além de criar um Ranking com base nos resultados que as lojas apresentaram, e é enviado à diretoria.

Um Resumo é muito simples e direto ao ponto, usado pela equipe de gerência de loja para saber os principais indicadores de cada loja e permitir comparar entre diferentes lojas quais metas aquela loja conseguiu cumprir naquele dia ou não.

O Ranking mostra de forma mais nítida quais lojas estão se destacando e quais estão a baixo da média, dessa forma, facilitando enxergar quais lojas estão apresentando os indicadores acima das metas estipuladas.

O Objetivo é conseguir criar um processo da forma mais automática possível para calcular o Resumo de cada loja e enviar um email para o gerente de cada loja com o seu Resumo no corpo do e-mail e também o arquivo completo com os dados da sua respectiva loja em anexo.


### O que será automatizado?

- Criação de pasta e arquivo de dados para cada loja
- Calculo para metas e indicadores
- Criação do Ranking
- Criação e envio do email

#### O que é neccesário?
Foi necessário a utilização de base de dados com as informções de vendas (produto, valor, quantidade vendida, data da venda) e da loja (nome da loja, nome e email do gerente do gerente)