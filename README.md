# Gerador de Relatórios em Excel

Sistema Flask que gera planilhas profissionais com análise financeira, gráficos e recomendações automáticas. Usuário informa produto, valores e vendas; sistema calcula tudo e gera Excel completo com OpenPyXL.

## Como funciona?

1. O usuário preenche um formulário com:
   - Nome do produto
   - Mês
   - Valor de venda
   - Custo de produção
   - Quantidade vendida
   - Estoque disponível

2. O sistema calcula automaticamente:
   - Lucro por unidade
   - Receita total
   - Custo total
   - Lucro total
   - Margem de lucro (%)
   - Estoque que sobrou

3. O programa gera um arquivo Excel bonitinho com:
   - Todos os valores organizados
   - Gráficos (pizza e barras)
   - Cores (verde pra lucro, vermelho pra custo)
   - Análise se o negócio está indo bem ou não

## Como instalar e rodar?

1. Instale o Python no seu computador (versão 3.7 ou superior)

2. Abra o terminal e instale as bibliotecas:
pip install flask openpyxl


3. Execute o arquivo:
python app.py


4. Abra o navegador e acesse:
http://localhost:5000


## Tecnologias usadas

- **Python** - Linguagem principal
- **Flask** - Cria o site e as rotas
- **OpenPyXL** - Biblioteca que manipula o Excel
- **HTML/CSS** - Interface do formulário

## Estrutura do projeto
gerador-excel/
│
├── app.py # Código principal
├── README.md # Documentação
│
└── templates/
└── index.html # Página do formulário


## Exemplo prático

**Se você colocar:**
- Produto: Camiseta
- Valor: R$ 50,00
- Custo: R$ 30,00
- Vendeu: 100 unidades
- Estoque: 150 unidades

**O Excel vai mostrar:**
- Receita: R$ 5.000,00
- Lucro total: R$ 2.000,00
- Margem: 40%
- "Excelente margem de lucro! ⭐"

## Licença

MIT - Pode usar à vontade :)
