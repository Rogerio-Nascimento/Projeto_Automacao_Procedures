# Projeto de Automacao de Procedures

Uma automação em que ele gera um arquivo .CSV para chamar as procedures de geração de arquivo .txt no banco de dados SQL Server. Essas procedures são amplamente utilizadas em projetos tributários que envolvem grandes volumes de dados, como CNPJs e períodos.


## Como funciona?

É um executável no qual o usuário irá inserir as informações para que, no final, seja gerado um arquivo .txt com essas informações, evitando retrabalho.

### Passando o caminho

Primeiramente, o usuário irá fornecer o SERVIDOR, BANCO DE DADOS SQL e a PROCEDURE.

![image](https://github.com/Rogerio-Nascimento/Projeto_Automacao_Procedures/assets/87660080/ef0fdf0f-9edb-4abd-b48e-d1acee540064)

Logo em seguida, o usuário irá inserir o período, por exemplo:

2018 - 2024 (semelhante à imagem abaixo):

![image](https://github.com/Rogerio-Nascimento/Projeto_Automacao_Procedures/assets/87660080/c4ba2675-e02c-47f7-b86a-5352f5702c80)


Agora, o usuário irá inserir o parâmetro (dependendo do tipo de projeto) e, em seguida, clicar no botão 'Pasta de Destino' para salvar o arquivo final no formato .CSV.

![image](https://github.com/Rogerio-Nascimento/Projeto_Automacao_Procedures/assets/87660080/61176c62-05d9-4d4f-a51e-28e8b075b4e6)

Em seguida, ele irá marcar as opções para incluir a filial como parâmetro, excluir se existir e incluir cabeçalho. Por fim, clicará em 'Gerar .CSV'.

![image](https://github.com/Rogerio-Nascimento/Projeto_Automacao_Procedures/assets/87660080/dd26855c-655c-4981-82f1-6590300957e7)
