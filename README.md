# Inserindo gráfico do matplotlib no excel

### É necessário a instalação de duas bibliotecas
* pip install openpyxl
* pip install matplotlib

As inserções dos dados nas planilhas do excel poderiam ser através de uma lista de valores:
~~~py
list_values = [10, 20, 30, 40]
for l in range(0, 4):
    ws1.cell(l + 1, 1, list_values[l])
    
~~~

Podemos adicionar a imagem para cada tabela
~~~py
ws1.add_image('fig1.png', 'C1')
~~~

