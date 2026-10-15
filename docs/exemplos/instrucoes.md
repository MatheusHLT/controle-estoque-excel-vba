# Instruções de Utilização

Este documento explica como testar a planilha demonstrativa de controle de estoque, desenvolvida com Microsoft Excel, VBA e fórmulas.

## 1. Requisitos

- Microsoft Excel para Windows com suporte a macros VBA.
- O arquivo `controle-estoque-demonstracao.xlsm`, disponível nesta pasta do repositório.

**Segurança:** habilite macros somente se tiver confiança na origem do arquivo e tiver verificado o código. Não desative globalmente as proteções de segurança do Excel para executar a planilha.

## 2. Abrir a planilha

1. Baixe o arquivo `controle-estoque-demonstracao.xlsm` do repositório.
2. Abra-o no aplicativo Microsoft Excel para Windows.
3. Se o Excel apresentar avisos de segurança, leia-os antes de decidir se permite o conteúdo ativo. As macros são necessárias para executar os botões de registro.
4. Utilize a versão demonstrativa com os dados fictícios fornecidos; evite inserir informações pessoais ou dados reais de uma organização.

## 3. Organização das abas

- **Estoque:** apresenta os materiais e as quantidades disponíveis, além das áreas para registrar entradas e retiradas.
- **Registro:** contém o histórico das movimentações de materiais dessa categoria.
- **EstoqueLampada:** apresenta o controle da categoria de lâmpadas e a área para registrar suas movimentações.
- **RegistroLampada:** contém o histórico das movimentações de lâmpadas.

## 4. Registrar uma entrada

1. Abra a aba correspondente ao tipo de material que deseja movimentar.
2. Na área de entrada, preencha o nome do responsável, a data e as quantidades dos materiais recebidos.
3. Clique no botão **REGISTRAR ENTRADA**.
4. Confira se as quantidades disponíveis foram atualizadas e se a movimentação aparece na aba de registro correspondente.

## 5. Registrar uma retirada

1. Abra a aba correspondente ao tipo de material que deseja movimentar.
2. Na área de retirada, preencha o nome do responsável, a data e as quantidades dos materiais retirados.
3. Clique no botão **REGISTRAR RETIRADA**.
4. Confira se as quantidades disponíveis foram atualizadas e se a movimentação aparece na aba de registro correspondente.

As macros de retirada verificam as quantidades informadas e a disponibilidade no estoque, impedindo retiradas superiores ao saldo disponível.

## 6. Validações e alertas

A planilha combina validação de dados do Excel com verificações implementadas em VBA. Os campos configurados com validação podem restringir os valores aceitos. As macros também verificam condições necessárias para registrar entradas e retiradas.

Fórmulas do Excel sinalizam materiais que precisam de reposição, conforme os limites configurados na planilha.

## 7. Consultar o histórico

Após registrar uma operação, abra a aba **Registro** ou **RegistroLampada**, de acordo com a categoria utilizada. O histórico apresenta informações da movimentação, como responsável, data, tipo de operação e quantidades registradas.

## 8. Observações

- Utilize apenas os dados fictícios da versão demonstrativa ao testar o projeto.
- Confira o estoque e o histórico após cada operação.
- Evite alterar, mover ou excluir células e botões utilizados pelas macros, pois isso pode afetar o funcionamento da planilha.
- Este projeto tem finalidade demonstrativa e não deve ser utilizado para controlar estoque real sem testes e validações adicionais.
