# Projeto Automação de Indicadores em Python 
#### Desafio por HashtagTreinamentos
## DESCRIÇÃO DO PROJETO:
#### Imagine que você trabalha em uma grande rede de loja roupas com 25 lojas espalhadas pelo Brasil. Todo dia pela manhã a equipe de análise de dados calcula os One Pages e envia para o gerente de cada loja o OnePage correspondente, bem como todas as informações usadas no cálculo dos indicadores. Um OnePage é um resumo simples e direto, usado pela equipe de gerência para saber os principais indicadores de cada loja e permitir em uma página tanto a comparação entre diferentes lojas, quanto quais indicadores aquela loja conseguiu cumprir naquele dia ou não.
### Objetivo:
 Criar um processo automático para calcular o OnePage de cada loja e enviar um email para o gerente de cada loja com o OnePage correspondente e também o arquivo completo com os dados da sua respectiva loja em formato .xlsx. Além disso, enviar ainda um e-mail para a diretoria (informações estão no arquivo Emails.xlsx) com 2 rankings das lojas em anexo, 1 ranking diário e outro ranking anual. Ademais, no corpo do e-mail, é necessário ressaltar qual foi a melhor e a pior loja do dia e também a melhor e pior loja do ano. O ranking de uma loja é dado pelo faturamento dela.
### Arquivos e Informações Importantes

- Arquivo Emails.xlsx com o nome, a loja e o e-mail de cada gerente. Obs: Substituir a coluna de e-mail de cada gerente por um e-mail seu, para você poder testar o resultado

- Arquivo Vendas.xlsx com as vendas de todas as lojas. Obs: Cada gerente só recebe o OnePage e um arquivo em excel em anexo com as vendas da SUA loja. 

- Arquivo Lojas.csv com o nome de cada Loja

- As planilhas de cada loja são salvas dentro da pasta da loja com a data da planilha, a fim de criar um histórico de backup. Obs: criar a pasta "Backup Arquivos Lojas" antes.

### Indicadores Do OnePage

- Faturamento -> Meta Ano: 1.650.000 / Meta Dia: 1000
- Diversidade de Produtos (quantos produtos diferentes foram vendidos naquele período) -> Meta Ano: 120 / Meta Dia: 4
- Ticket Médio por Venda -> Meta Ano: 500 / Meta Dia: 500

Obs: Cada indicador é calculado no dia e no ano. O indicador do dia refere-se ao último dia disponível na planilha de Vendas (a data mais recente)

### Descrição do trabalho realizado pelo código:

- Abre os arquivos com os dados, cria uma tabela para cada loja e define o dia do indicador.
- Salva as tabelas (planilhas) na pasta de Backup.
- Calcula os indicadores (Faturamento, Diversidade de Produtos, Ticket Médio) tendo como base as metas já definidas.
- Envia o email com o OnePage para cada gerente, anexando uma planinha excel com os dados da loja específica.
- Cria um ranking para a diretoria, tendo em vista os faturamentos anual e diário.
- Envia E-mail para a diretoria, anexando os rankings e inserindo uma mensagem para o corpo do E-mail.

### Bibliotecas Utilizadas:
- pandas
- pathlib
- win32com.client
- time

## Anexos
### E-mail para a gerência
![image](https://user-images.githubusercontent.com/118035572/216399601-88b3f4e4-04bd-4bb2-a657-5e5b5bc3ef11.png)
### E-mail para a Diretoria
![image](https://user-images.githubusercontent.com/118035572/216398120-a7c07555-b0ac-4abf-a269-fa994ab5f775.png)



